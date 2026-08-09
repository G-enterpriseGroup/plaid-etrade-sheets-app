import os, json, base64, math, time
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
import pandas as pd
import yfinance as yf
from google.oauth2 import service_account
from googleapiclient.discovery import build

SID=os.getenv('GOOGLE_SHEET_ID','1sgNWMAZEIdOBargwH8ILs_oB-1HQzUoVssaqnsujY2g')
PUT='Put Money Flow'; CACHE='Option Quote Cache'; PICKER='YF Option Picker'
FIRST=73; LAST=112; ET=ZoneInfo('America/New_York'); TTL=18
SCOPE='https://www.googleapis.com/auth/spreadsheets'

def log(x): print(datetime.now(ET).strftime('%H:%M:%S'),x,flush=True)
def col(n):
    s=''
    while n: n,r=divmod(n-1,26); s=chr(65+r)+s
    return s
def fnum(x):
    try:
        n=float(str(x or '').replace('$','').replace(',',''))
        return n if math.isfinite(n) else None
    except: return None
def expdate(x):
    x=str(x or '').strip()
    if not x:return ''
    for fmt in ('%Y-%m-%d','%m/%d/%Y','%m/%d/%y'):
        try:return datetime.strptime(x,fmt).date().isoformat()
        except:pass
    try:return pd.to_datetime(x).date().isoformat()
    except:return x
def serial(x): return (datetime.strptime(x,'%Y-%m-%d').date()-date(1899,12,30)).days
def now(): return datetime.now(ET).replace(microsecond=0).isoformat()

def auth_info():
    raw=os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON','').strip()
    if not raw: raise RuntimeError('Missing GOOGLE_SERVICE_ACCOUNT_JSON GitHub secret')
    try:return json.loads(raw)
    except:
        try:return json.loads(base64.b64decode(raw).decode())
        except:raise RuntimeError('GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON/base64 JSON')
def svc():
    c=service_account.Credentials.from_service_account_info(auth_info(),scopes=[SCOPE])
    return build('sheets','v4',credentials=c,cache_discovery=False)
def meta(s):
    r=s.spreadsheets().get(spreadsheetId=SID,fields='properties.title,sheets.properties').execute()
    return r.get('properties',{}),{x['properties']['title']:x['properties'] for x in r['sheets']}
def vals(s,ranges):
    r=s.spreadsheets().values().batchGet(spreadsheetId=SID,ranges=ranges,valueRenderOption='FORMATTED_VALUE').execute()
    return [x.get('values',[]) for x in r.get('valueRanges',[])]
def write(s,data,user=True):
    s.spreadsheets().values().batchUpdate(spreadsheetId=SID,body={'valueInputOption':'USER_ENTERED' if user else 'RAW','data':data}).execute()
def clear(s,r): s.spreadsheets().values().clear(spreadsheetId=SID,range=r,body={}).execute()

def ensure_sheet(s,name,rows,cols,hidden=True):
    _,m=meta(s)
    if name not in m:
        s.spreadsheets().batchUpdate(spreadsheetId=SID,body={'requests':[{'addSheet':{'properties':{'title':name,'hidden':hidden,'gridProperties':{'rowCount':rows,'columnCount':cols,'frozenRowCount':1}}}}]}).execute()
    _,m=meta(s); p=m[name]; g=p.get('gridProperties',{}); req=[]
    if g.get('rowCount',0)<rows or g.get('columnCount',0)<cols:
        req.append({'updateSheetProperties':{'properties':{'sheetId':p['sheetId'],'gridProperties':{'rowCount':max(rows,g.get('rowCount',0)),'columnCount':max(cols,g.get('columnCount',0))}},'fields':'gridProperties(rowCount,columnCount)'}})
    if hidden and not p.get('hidden',False): req.append({'updateSheetProperties':{'properties':{'sheetId':p['sheetId'],'hidden':True},'fields':'hidden'}})
    if req:s.spreadsheets().batchUpdate(spreadsheetId=SID,body={'requests':req}).execute()
    return meta(s)[1][name]

def setup_sheet(s):
    pp=ensure_sheet(s,PICKER,20000,100,True); ensure_sheet(s,CACHE,500,7,True); _,m=meta(s); putid=m[PUT]['sheetId']; pid=pp['sheetId']
    writes=[{'range':f"'{PICKER}'!A1:F1",'values':[['Ticker','Expiration','Type','Strike','Contract Symbol','Refreshed At']]},
            {'range':f"'{CACHE}'!A1:G1",'values':[['Option Quote Key','Bid','Ask','Last','Mark','Source / Status','Updated At']]},
            {'range':f"'{PUT}'!A70",'values':[['AUTOMATED OPTION PICKER: type ticker → choose Yahoo/yfinance expiration → choose strike. GitHub Actions refreshes option marks automatically.']]},
            {'range':f"'{PUT}'!R72",'values':[['Option Quote Key']]}]
    req=[{'repeatCell':{'range':{'sheetId':pid,'startRowIndex':1,'endRowIndex':20000,'startColumnIndex':1,'endColumnIndex':2},'cell':{'userEnteredFormat':{'numberFormat':{'type':'DATE','pattern':'yyyy-mm-dd'}}},'fields':'userEnteredFormat.numberFormat'}}]
    for i,r in enumerate(range(FIRST,LAST+1)):
        ec=8+i*2; sc=ec+1; E=col(ec); K=col(sc)
        ef=f'=IFERROR(SORT(UNIQUE(FILTER($B$2:$B$20000,$A$2:$A$20000=\'{PUT}\'!$D{r},$B$2:$B$20000>=TODAY()))),"")'
        sf=f'=IFERROR(SORT(UNIQUE(FILTER($D$2:$D$20000,$A$2:$A$20000=\'{PUT}\'!$D{r},$B$2:$B$20000=\'{PUT}\'!$E{r},$C$2:$C$20000=IF(REGEXMATCH(\'{PUT}\'!$C{r},"CALL"),"CALL","PUT")))),"")'
        writes += [{'range':f"'{PICKER}'!{E}1:{E}2",'values':[[f'Row {r} expirations'],[ef]]},{'range':f"'{PICKER}'!{K}1:{K}2",'values':[[f'Row {r} strikes'],[sf]]}]
        R=f'=IF(OR(NOT(REGEXMATCH($C{r},"CALL|PUT|LEAPS")),$D{r}="",$E{r}="",$F{r}=""),"",UPPER($D{r})&":"&YEAR($E{r})&":"&MONTH($E{r})&":"&DAY($E{r})&":"&IF(REGEXMATCH($C{r},"CALL"),"CALL","PUT")&":"&TEXT($F{r},"0.###"))'
        N=f'=IF($C{r}="","",IF(OR($C{r}="STOCK",$C{r}="ETF"),"GOOGLEFINANCE",IFERROR(INDEX(\'{CACHE}\'!$F:$F,MATCH($R{r},\'{CACHE}\'!$A:$A,0)),"WAITING FOR YAHOO")))'
        Q=f'=IF($C{r}="","",IF(OR($C{r}="STOCK",$C{r}="ETF"),"Google Finance",IF($N{r}="YFINANCE","Yahoo/yfinance mark • auto refreshed","Waiting for Yahoo/yfinance refresh")))'
        writes += [{'range':f"'{PUT}'!N{r}",'values':[[N]]},{'range':f"'{PUT}'!R{r}",'values':[[R]]},{'range':f"'{PUT}'!S{r}",'values':[[Q]]}]
        req += [
          {'repeatCell':{'range':{'sheetId':pid,'startRowIndex':1,'endRowIndex':121,'startColumnIndex':ec-1,'endColumnIndex':ec},'cell':{'userEnteredFormat':{'numberFormat':{'type':'DATE','pattern':'yyyy-mm-dd'}}},'fields':'userEnteredFormat.numberFormat'}},
          {'setDataValidation':{'range':{'sheetId':putid,'startRowIndex':r-1,'endRowIndex':r,'startColumnIndex':4,'endColumnIndex':5},'rule':{'condition':{'type':'ONE_OF_RANGE','values':[{'userEnteredValue':f"='{PICKER}'!${E}$2:${E}$120"}]},'strict':True,'showCustomUi':True}}},
          {'setDataValidation':{'range':{'sheetId':putid,'startRowIndex':r-1,'endRowIndex':r,'startColumnIndex':5,'endColumnIndex':6},'rule':{'condition':{'type':'ONE_OF_RANGE','values':[{'userEnteredValue':f"='{PICKER}'!${K}$2:${K}$600"}]},'strict':True,'showCustomUi':True}}}]
    write(s,writes,True); s.spreadsheets().batchUpdate(spreadsheetId=SID,body={'requests':req}).execute()

def read_rows(s):
    b,p=vals(s,[f"'{PUT}'!A{FIRST}:S{LAST}",f"'{PUT}'!B26:B65"]); complete=[]; tickers=[]
    for i,x in enumerate(b):
        x=x+['']*(19-len(x)); typ=str(x[2]).upper(); t=str(x[3]).strip().upper()
        if not any(z in typ for z in ('CALL','PUT','LEAPS')):continue
        if t and t not in tickers:tickers.append(t)
        e=expdate(x[4]); k=fnum(x[5]); key=str(x[17]).strip()
        if t and e and k is not None and key:complete.append((FIRST+i,typ,t,e,k,key))
    for x in p:
        if x:
            t=str(x[0]).strip().upper()
            if t and t not in tickers:tickers.append(t)
    return complete,tickers

def retry(label,fn):
    err=None
    for n in range(3):
        try:return fn()
        except Exception as e:err=e; log(f'{label} retry {n+1}/3: {e}'); time.sleep(2*(n+1))
    raise RuntimeError(f'{label}: {err}')
class Yahoo:
    def __init__(self):self.t={};self.c={};self.e={}
    def T(self,t):
        if t not in self.t:self.t[t]=yf.Ticker(t)
        return self.t[t]
    def exps(self,t):
        if t not in self.e:self.e[t]=tuple(retry(t+' expirations',lambda: self.T(t).options or ()))
        return self.e[t]
    def chain(self,t,e):
        if (t,e) not in self.c:self.c[(t,e)]=retry(t+' '+e,lambda:self.T(t).option_chain(e))
        return self.c[(t,e)]

def old_universe(s):
    u=vals(s,[f"'{PICKER}'!A2:F20000"])[0]; out=[]
    for x in u:
        x=x+['']*(6-len(x)); t=str(x[0]).upper(); e=expdate(x[1]); k=fnum(x[3])
        if t and e and str(x[2]).upper() in ('CALL','PUT') and k is not None:out.append([t,e,str(x[2]).upper(),k,str(x[4]),str(x[5])])
    return out
def fresh(u,t):
    ds=[]
    for x in u:
        if x[0]==t:
            try:ds.append(datetime.fromisoformat(x[5]).astimezone(ET))
            except:pass
    return bool(ds) and datetime.now(ET)-max(ds)<timedelta(hours=TTL)
def refresh_universe(s,y,tickers):
    u=old_universe(s); changed=False
    for t in tickers:
        if fresh(u,t):continue
        stamp=now(); nr=[]; exps=list(y.exps(t))[:80]
        for e in exps:
            try:
                ch=y.chain(t,e)
                for typ,df in [('CALL',ch.calls),('PUT',ch.puts)]:
                    if df is None or df.empty:continue
                    for _,z in df.iterrows():
                        k=fnum(z.get('strike'))
                        if k is not None:nr.append([t,e,typ,k,str(z.get('contractSymbol') or ''),stamp])
            except Exception as er:log(f'{t} {e} picker skipped: {er}')
        if nr:u=[x for x in u if x[0]!=t]+nr;changed=True;log(f'{t}: cached {len(nr)} option contracts')
    if not changed:return
    u=sorted(u,key=lambda x:(x[0],x[1],x[2],x[3])); clear(s,f"'{PICKER}'!A2:F")
    rows=[[x[0],serial(x[1]),x[2],x[3],x[4],x[5]] for x in u]
    if rows:write(s,[{'range':f"'{PICKER}'!A2:F{len(rows)+1}",'values':rows}],False)

def quote(df,k):
    if df is None or df.empty or 'strike' not in df:return None
    st=pd.to_numeric(df['strike'],errors='coerce'); q=df[(st-k).abs()<1e-6]
    if q.empty:return None
    z=q.iloc[0]; b=fnum(z.get('bid'));a=fnum(z.get('ask'));l=fnum(z.get('lastPrice'));m=None
    if b and a and a>=b:m=(b+a)/2
    elif l and l>0:m=l
    elif b and b>0:m=b
    elif a and a>0:m=a
    return b,a,l,m
def refresh_quotes(s,y,rows):
    out=[]; stamp=now()
    for r,typ,t,e,k,key in rows:
        b=a=l=m=''; status='YFINANCE'
        try:
            ch=y.chain(t,e); q=quote(ch.calls if 'CALL' in typ else ch.puts,k)
            if q is None:status='NO YFINANCE CONTRACT'
            else:
                b,a,l,m=q; b='' if b is None else b;a='' if a is None else a;l='' if l is None else l;m='' if m is None else m
                if m=='':status='NO YFINANCE QUOTE'
        except Exception as er:status='YFINANCE ERROR';log(f'row {r} {key}: {er}')
        out.append([key,b,a,l,m,status,stamp])
    clear(s,f"'{CACHE}'!A2:G")
    if out:write(s,[{'range':f"'{CACHE}'!A2:G{len(out)+1}",'values':out}],False)
    log(f'updated {len(out)} selected option quote(s)')

def main():
    s=svc(); title,_=meta(s);log('connected: '+title.get('title',SID));setup_sheet(s)
    rows,tickers=read_rows(s);log(f'{len(rows)} fully-selected option row(s); picker tickers={tickers}')
    y=Yahoo();refresh_universe(s,y,tickers);refresh_quotes(s,y,rows);log('DONE')
if __name__=='__main__':main()
