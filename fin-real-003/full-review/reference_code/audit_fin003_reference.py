"""Independent pandas aggregation checks of the Decimal reference candidate."""
import json
from pathlib import Path
import pandas as pd

ROOT=Path(__file__).resolve().parent/'redesign/FIN-REAL-003'


def main():
    data=pd.read_csv(ROOT/'input_files/fdic_financials_VT_NH_ME_2022_2025.csv',dtype={'CERT':str,'REPDTE':str})
    institutions=pd.read_csv(ROOT/'input_files/fdic_institutions_VT_NH_ME.csv',dtype={'CERT':str,'NEWCERT':str})
    result=json.loads((ROOT/'reference_output/results.json').read_text())
    checks=[]
    def check(label,actual,expected,tol=1e-8):
        checks.append(dict(check=label,passed=abs(float(actual)-float(expected))<=tol,actual=float(actual),expected=float(expected)))
    for k,v in dict(financial_rows=len(data),quarters=data.REPDTE.nunique(),banks=data.CERT.nunique(),
                    matched_rows=data.merge(institutions[['CERT']],on='CERT',validate='many_to_one').shape[0]).items():
        check(k,result['counts'][k],v,0)
    comparable={r['cert'] for r in result['coverage'] if r['comparable']}
    fixed={r['cert'] for r in result['coverage'] if not r['missing']}
    for r in result['aggregates']:
        rows=data[data.REPDTE==r['date']]
        if r['state']!='ALL': rows=rows[rows.STALP==r['state']]
        if r['scope']=='continuous': rows=rows[rows.CERT.isin(fixed)]
        if r['scope']=='continuous_conventional': rows=rows[rows.CERT.isin(comparable)]
        positive=rows[rows.RBCRWAJ>0]
        expected=dict(banks=len(rows),asset=rows.ASSET.sum(),dep=rows.DEP.sum(),
            roa_asset_proxy=(rows.ROA*rows.ASSET).sum()/rows.ASSET.sum(),
            nonaccrual_assets_pct=100*rows.NAASSET.sum()/rows.ASSET.sum(),
            nco_loan_proxy=100*(rows.NTLNLS*12/rows.REPDTE.str[4:6].astype(int)).sum()/rows.LNLSGR.sum(),
            nco_asset_proxy=100*(rows.NTLNLS*12/rows.REPDTE.str[4:6].astype(int)).sum()/rows.ASSET.sum(),
            rbc_asset_proxy=(positive.RBCRWAJ*positive.ASSET).sum()/positive.ASSET.sum(),
            equity_asset_pct=100*rows.EQ.sum()/rows.ASSET.sum(),
            loan_deposit_pct=100*rows.LNLSNET.sum()/rows.DEP.sum(),
            real_estate_loan_pct=100*rows.LNRE.sum()/rows.LNLSGR.sum(),
            nim_asset_proxy=(rows.NIMY*rows.ASSET).sum()/rows.ASSET.sum())
        for k,v in expected.items(): check('/'.join([r['date'],r['state'],r['scope'],k]),r[k],v)
    for r in result['decomposition']:
        values=[x for x in result['aggregates'] if x['state']=='ALL' and x['scope']=='all' and x['date'] in (result['dates'][0],result['dates'][-1])]
        values.sort(key=lambda x:x['date'])
        for position,label in [(0,'start'),(1,'end')]: check(r['metric']+'/'+label,r[label],values[position][r['metric']])
        check(r['metric']+'/sum',sum(float(r[k]) for k in ['within_pp','reweight_pp','entry_pp','exit_pp']),float(r['end'])-float(r['start']))
    out=dict(checks=checks,failed=sum(not r['passed'] for r in checks),
             method='Independent pandas float aggregations vs Decimal reference; tolerance 1e-8 in stored source units.',
             scope='Numerical checks only; not independent business judgement or Office rendering approval.',gold_approved=False)
    (ROOT/'reference_output/independent_numeric_audit.json').write_text(json.dumps(out,ensure_ascii=False,indent=2))
    print(json.dumps(dict(checks=len(checks),failed=out['failed'],gold_approved=False)))
    if out['failed']: raise SystemExit(1)


if __name__=='__main__': main()
