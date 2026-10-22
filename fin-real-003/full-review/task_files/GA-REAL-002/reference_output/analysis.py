"""Reference computation from unmodified public sources, not a solver answer leak."""
import argparse
from collections import Counter, defaultdict
import csv
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
import json
import math
from pathlib import Path
import re
import sys

sys.dont_write_bytecode = True


def records(path, encoding='utf-8-sig'):
    with path.open(encoding=encoding, newline='') as stream:
        return [dict(row, _row=i) for i, row in enumerate(csv.DictReader(stream), 2)]


def number(value):
    try:
        result = Decimal(str(value).strip())
        return result if result.is_finite() else None
    except (InvalidOperation, ValueError):
        return None


def ratio(a, b):
    return a / b * 100 if a is not None and b is not None and b > 0 else None


def difference(a, b):
    return a-b if a is not None and b is not None else None


def total(rows, key):
    values = [number(r.get(key)) for r in rows]
    if any(v is None for v in values):
        raise ValueError('Missing numeric field: '+key)
    return sum(values, Decimal(0))


def go(inp):
    payments = records(inp/'Chicago_Payments_2024.csv')
    contracts = records(inp/'Chicago_Contracts_2019plus.csv')
    versions = defaultdict(list)
    for row in contracts:
        versions[row['purchase_order_contract_number']].append(row)
    def select(items, cutoff):
        eligible = [r for r in items if r['approval_date'] and r['approval_date'][:10] <= cutoff]
        if not eligible:
            return []
        key = max((r['approval_date'], int(r['revision_number'])) for r in eligible)
        latest = [r for r in eligible if (r['approval_date'], int(r['revision_number'])) == key]
        unique = {}
        for r in latest:
            identity = tuple(sorted((k,v) for k,v in r.items() if k != '_row'))
            unique.setdefault(identity, r)
        return list(unique.values())
    asof = {key:select(rows, '2024-12-31') for key,rows in versions.items()}
    alltime = {key:select(rows, '9999-12-31') for key,rows in versions.items()}
    groups = defaultdict(list)
    bad_dates = []
    for row in payments:
        when = row['check_date']
        if when == '2024':
            grain = 'annual'
        else:
            try:
                parsed = datetime.strptime(when, '%m/%d/%Y')
                grain = 'detail' if parsed.year == 2024 else 'other'
            except ValueError:
                grain = 'other'
        row = dict(row, grain=grain)
        if grain == 'other':
            bad_dates.append(row)
        groups[grain, row['contract_number']].append(row)
    summaries = []
    versions_changed = []
    matched_rows = []
    unmatched_rows = []
    exceptions = []
    for (grain, key), rows in groups.items():
        direct = not key.strip() or key.strip().upper() == 'DV'
        selected = asof.get(key, []) if not direct else []
        status = 'direct' if direct else 'matched' if len(selected) == 1 else 'unmatched' if not selected else 'conflict'
        record = dict(grain=grain, contract=key, payment_rows=[r['_row'] for r in rows],
                      payment_count=len(rows), amount=total(rows, 'amount'), status=status,
                      selected_contract_rows=[r['_row'] for r in selected])
        if len(selected) == 1:
            chosen = selected[0]
            record.update(vendor=chosen['vendor_name'], department=chosen['department'],
                          approval_date=chosen['approval_date'], revision=chosen['revision_number'],
                          award=number(chosen['award_amount']))
            record['paid_award_pct'] = ratio(record['amount'], record['award'])
            record['next_step'] = '调取完整合同及付款支持文件；年度付款和授予额不等同，当前比率不能直接证明超支或违法。'
            matched_rows.append(record)
            if record['award'] is None or record['award'] <= 0:
                exceptions.append(record)
        elif status in ('unmatched','conflict'):
            record['next_step'] = '补查截点前完整合同、修订链和归属；不凭供应商近似名称分配合同。'
            unmatched_rows.append(record)
        summaries.append(record)
    for key, original in asof.items():
        newer = alltime[key]
        if [r['_row'] for r in original] != [r['_row'] for r in newer]:
            related = [r for r in summaries if r['contract'] == key]
            versions_changed.append(dict(contract=key, asof_rows=[r['_row'] for r in original],
                                         latest_rows=[r['_row'] for r in newer],
                                         annual_amount=sum((r['amount'] for r in related if r['grain']=='annual'),Decimal(0)),
                                         detail_amount=sum((r['amount'] for r in related if r['grain']=='detail'),Decimal(0))))
    conservation = []
    for grain in sorted({k[0] for k in groups}):
        subset = [r for r in summaries if r['grain']==grain]
        conservation.append(dict(grain=grain, raw_rows=sum(r['payment_count'] for r in subset),
                                 raw_amount=sum((r['amount'] for r in subset),Decimal(0)),
                                 matched_amount=sum((r['amount'] for r in subset if r['status']=='matched'),Decimal(0)),
                                 unmatched_amount=sum((r['amount'] for r in subset if r['status']=='unmatched'),Decimal(0)),
                                 conflict_amount=sum((r['amount'] for r in subset if r['status']=='conflict'),Decimal(0)),
                                 direct_amount=sum((r['amount'] for r in subset if r['status']=='direct'),Decimal(0))))
    annual_keys = {k for grain,k in groups if grain=='annual'}
    detail_keys = {k for grain,k in groups if grain=='detail'}
    top_unmatched = sorted([r for r in unmatched_rows if r['grain']=='annual'],key=lambda r:(-abs(r['amount']),r['contract']))[:20]
    top_ratios = sorted([r for r in matched_rows if r['grain']=='annual' and r['paid_award_pct'] is not None],key=lambda r:(-r['paid_award_pct'],r['contract']))[:20]
    versions_changed.sort(key=lambda r:(-abs(r['annual_amount']),r['contract']))
    return dict(summary=dict(payment_rows=len(payments),contract_version_rows=len(contracts),
                             contract_numbers=len(versions),negative_payment_rows=sum(number(r['amount'])<0 for r in payments),
                             payment_unique_vouchers=len({r['voucher_number'] for r in payments}),
                             duplicate_contract_revision_keys=sum(v-1 for v in Counter((r['purchase_order_contract_number'],r['revision_number']) for r in contracts).values()),
                             grain_overlap=sorted(annual_keys & detail_keys),
                             future_changed_contracts=len(versions_changed),
                             paid_future_changed_contracts=sum(bool(r['annual_amount'] or r['detail_amount']) for r in versions_changed),
                             total_actual_spend=None),
                conservation=conservation, versions_changed=versions_changed,
                top_unmatched=top_unmatched,top_ratios=top_ratios,
                exceptions=exceptions,bad_dates=bad_dates,contract_payments=summaries,
                negative_payments=[r for r in payments if number(r['amount'])<0])


def fin(inp):
    raw = records(inp/'fdic_financials_VT_NH_ME_2022_2025.csv')
    institutions = records(inp/'fdic_institutions_VT_NH_ME.csv')
    index = defaultdict(list)
    for row in institutions:
        index[row['CERT']].append(row)
    keys = Counter((r['CERT'],r['REPDTE']) for r in raw)
    if any(v!=1 for v in keys.values()):
        raise ValueError('Duplicate bank-quarter requires individual resolution')
    quarters = sorted({r['REPDTE'] for r in raw})
    panel = []
    for row in raw:
        nums = {k:number(v) for k,v in row.items() if k not in ('NAMEFULL','STALP','_row','ID')}
        p = dict(cert=row['CERT'], date=row['REPDTE'], state=row['STALP'], name=row['NAMEFULL'],
                 source_row=row['_row'], institution_rows=[r['_row'] for r in index[row['CERT']]],
                 raw={k:nums[k] for k in ('ASSET','DEP','EQ','LNLSGR','LNLSNET','LNRE','P9ASSET','NAASSET','NCLNLSR','ROA','NIMY','RBCRWAJ')})
        p.update(asset=nums['ASSET'],dep=nums['DEP'],eq_asset_pct=ratio(nums['EQ'],nums['ASSET']),
                 loan_deposit_pct=ratio(nums['LNLSNET'],nums['DEP']),
                 real_estate_loan_pct=ratio(nums['LNRE'],nums['LNLSGR']),
                 quality_numerator=nums['P9ASSET']+nums['NAASSET'])
        p['quality_proxy_pct'] = ratio(p['quality_numerator'],p['asset'])
        panel.append(p)
    observations = defaultdict(list)
    for row in panel:
        observations[row['cert']].append(row)
    fixed = {k for k,rows in observations.items() if {r['date'] for r in rows}==set(quarters)}
    aggregates = []
    for quarter in quarters:
        for state in ['ME','NH','VT','ALL']:
            for mode in ['dynamic','fixed']:
                group = [p for p in panel if p['date']==quarter and (state=='ALL' or p['state']==state) and (mode=='dynamic' or p['cert'] in fixed)]
                assets = sum((r['asset'] for r in group),Decimal(0))
                numerator = sum((r['quality_numerator'] for r in group),Decimal(0))
                aggregates.append(dict(date=quarter,state=state,sample=mode,banks=len(group),asset=assets,
                                       deposits=sum((r['dep'] for r in group),Decimal(0)),
                                       quality_numerator=numerator,pooled_pct=ratio(numerator,assets),
                                       equal_pct=sum((r['quality_proxy_pct'] for r in group),Decimal(0))/len(group) if group else None,
                                       weighted_pct=sum((r['quality_proxy_pct']*r['asset'] for r in group),Decimal(0))/assets if assets>0 else None))
    endpoints = []
    for q in [quarters[0],quarters[-1]]:
        all_rows = [r for r in panel if r['date']==q]
        removed = [r for r in all_rows if len(index[r['cert']])!=1 or index[r['cert']][0]['ACTIVE']!='1']
        active = [r for r in all_rows if r not in removed]
        endpoints.append(dict(date=q,all_banks=len(all_rows),active_banks=len(active),
                              removed_certs=[r['cert'] for r in removed],
                              excluded_asset=sum((r['asset'] for r in removed),Decimal(0)),
                              all_proxy_pct=ratio(sum((r['quality_numerator'] for r in all_rows),Decimal(0)),sum((r['asset'] for r in all_rows),Decimal(0))),
                              active_proxy_pct=ratio(sum((r['quality_numerator'] for r in active),Decimal(0)),sum((r['asset'] for r in active),Decimal(0)))))
    last = {r['cert']:r for r in panel if r['date']=='20250630'}
    previous = {r['cert']:r for r in panel if r['date']=='20240630'}
    ranking = []
    for cert in last.keys() & previous.keys():
        now, before = last[cert],previous[cert]
        ranking.append(dict(cert=cert,name=now['name'],source_rows=[before['source_row'],now['source_row']],
                            quality_pct=now['quality_proxy_pct'],yoy_pp=now['quality_proxy_pct']-before['quality_proxy_pct'],
                            eq_asset_change_pp=now['eq_asset_pct']-before['eq_asset_pct'],
                            real_estate_change_pp=difference(now['real_estate_loan_pct'],before['real_estate_loan_pct']),
                            rbcrwaj_change_pp=now['raw']['RBCRWAJ']-before['raw']['RBCRWAJ'],
                            previous=before,current=now))
    ranking.sort(key=lambda r:(-r['yoy_pp'],-r['quality_pct'],int(r['cert'])))
    mismatches = [dict(cert=r['cert'],source_row=r['source_row'],institution_rows=r['institution_rows'],
                       issue='missing_or_duplicate_institution' if len(index[r['cert']])!=1 else 'name_or_state_difference')
                  for r in panel if len(index[r['cert']])!=1 or index[r['cert']][0]['NAME']!=r['name'] or index[r['cert']][0]['STALP']!=r['state']]
    return dict(summary=dict(observations=len(panel),banks=len(observations),quarters=quarters,fixed_banks=len(fixed)),
                panel=panel, aggregates=aggregates, endpoints=endpoints, top10=ranking[:10],
                coverage=[dict(cert=k,quarters=[r['date'] for r in rows],missing=sorted(set(quarters)-{r['date'] for r in rows}),fixed=k in fixed) for k,rows in sorted(observations.items(),key=lambda kv:int(kv[0]))],
                institution_exceptions=mismatches)


def parse_date(value):
    value = (value or '').strip()
    if not value or value.lower()=='until canceled':
        return None
    for fmt in ('%Y-%m-%d','%m/%d/%Y'):
        try:
            return datetime.strptime(value[:10],fmt).date()
        except ValueError:
            pass
    raise ValueError('Unrecognized date: '+value)


def law(inp):
    general = records(inp/'LNI_Contractor_License_General_BR.csv')
    bonds = records(inp/'LNI_Contractor_License_Bond_BR.csv')
    insurance = records(inp/'LNI_Contractor_License_Insurance_BR.csv')
    dates = [date(2026,10,12),date(2026,11,11)]
    grouped = {kind:defaultdict(list) for kind in ('bond','insurance')}
    policies = []
    for kind, rows in [('bond',bonds),('insurance',insurance)]:
        for row in rows:
            if kind=='bond':
                field_names = ['bondeffectivedate','bondexpirationdate','bondcanceldate','bondimpaireddate']
                amount = number(row['bondamt'])
            else:
                field_names = ['effectivedate','expirationdate','canceldate']
                amount = number(row['insuranceamt'])
            parsed = {}
            issues = []
            for key in field_names:
                try:
                    parsed[key] = parse_date(row[key])
                except ValueError:
                    parsed[key] = None
                    issues.append(key+':unparsed')
            start_key,end_key = field_names[:2]
            if parsed[start_key] is None:
                issues.append(start_key+':missing')
            if parsed[end_key] is None and row[end_key].strip().lower()!='until canceled':
                issues.append(end_key+':missing')
            impaired = kind=='bond' and row.get('bondimpaired','').strip().upper() not in ('','N','NO','FALSE','0')
            if impaired and parsed['bondimpaireddate'] is None:
                issues.append('impaired_without_date')
            ends = [v for k,v in parsed.items() if k!=start_key and v is not None]
            end = min(ends) if ends else None
            valid = {}
            for day in dates:
                if parsed[start_key] is not None and parsed[start_key]>day or end is not None and end<=day:
                    value=False
                elif issues:
                    value=None
                else:
                    value=True
                valid[day.isoformat()] = value
            record = dict(kind=kind,license=row['contractorlicensenumber'],ubi=row['ubi'],source_row=row['_row'],
                          amount=amount,raw_dates={key:row[key] for key in field_names},
                          start=parsed[start_key].isoformat() if parsed[start_key] else None,
                          end=end.isoformat() if end else None,issues=issues,valid=valid,
                          policy_reference=row['bondaccountid'] if kind=='bond' else row['insurancepolicyno'])
            policies.append(record)
            grouped[kind][record['license']].append(record)
    licenses = []
    ubi_licenses = defaultdict(set)
    general_keys = Counter(r['contractorlicensenumber'] for r in general)
    exceptions = []
    for row in general:
        key = row['contractorlicensenumber']
        ubi_licenses[row['ubi']].add(key)
        evidence = {kind:grouped[kind][key] for kind in grouped}
        legal_general = row['specialtycode1desc'].strip().upper()=='GENERAL'
        minimum = Decimal(30000 if legal_general else 15000) if row['contractorlicensetypecode']=='CC' and row['specialtycode1desc'] else None
        status = {}
        for day in dates:
            ds = day.isoformat()
            bv = [r for r in evidence['bond'] if r['valid'][ds] is True]
            iv = [r for r in evidence['insurance'] if r['valid'][ds] is True]
            status[ds] = dict(bond_valid_rows=[r['source_row'] for r in bv],insurance_valid_rows=[r['source_row'] for r in iv],
                              bond_amount_supported=any(r['amount'] is not None and r['amount']>=minimum for r in bv) if minimum is not None else None,
                              insurance_form_proven=False,
                              unresolved_bond_rows=[r['source_row'] for r in evidence['bond'] if r['valid'][ds] is None],
                              unresolved_insurance_rows=[r['source_row'] for r in evidence['insurance'] if r['valid'][ds] is None])
        active = row['contractorlicensestatus']=='ACTIVE'
        now,later = [status[d.isoformat()] for d in dates]
        current_gap = active and not now['bond_valid_rows'] and not now['insurance_valid_rows']
        gap_days={}
        for kind in ['bond','insurance']:
            gaps=[]
            for offset in range(1,31):
                day=dates[0]+timedelta(days=offset)
                supported=any(r['start'] and r['start']<=day.isoformat() and (not r['end'] or day.isoformat()<r['end']) and not r['issues'] for r in evidence[kind])
                if not supported:
                    gaps.append(day.isoformat())
            gap_days[kind]=gaps
        renewal_gap = active and ((bool(now['bond_valid_rows']) and bool(gap_days['bond'])) or (bool(now['insurance_valid_rows']) and bool(gap_days['insurance'])))
        record = dict(license=key,business=row['businessname'],ubi=row['ubi'],source_row=row['_row'],
                      reported_status=row['contractorlicensestatus'],specialty=row['specialtycode1desc'],
                      registration_effective=row['licenseeffectivedate'],registration_expiration=row['licenseexpirationdate'],
                      bond_minimum=minimum,status=status,license_type=row['contractorlicensetypecode'],
                      legal_scope='construction_registration_scope' if row['contractorlicensetypecode']=='CC' else 'not_established_by_this_package',
                      future_coverage_gap_days=gap_days,
                      bond_rows=[r['source_row'] for r in evidence['bond']],insurance_rows=[r['source_row'] for r in evidence['insurance']],
                      active_current_both_gap=current_gap,active_renewal_gap=renewal_gap)
        licenses.append(record)
        for kind,items in evidence.items():
            for item in items:
                if item['ubi'] != row['ubi']:
                    exceptions.append(dict(kind=kind,license=key,general_row=row['_row'],policy_row=item['source_row'],issue='UBI mismatch'))
    orphans = [r for r in policies if r['license'] not in general_keys]
    legal_matrix=[
        dict(clause='18.27.010(1),(5),(12)',pdf_pages='2-3',topic='适用范围与一般/专业承包商',finding='建筑工作范围和身份是法律适用依据，不能把所有登记类型视作同一类别。'),
        dict(clause='18.27.010(15)',pdf_pages='3-4',topic='登记届满30天例外',finding='尚须维持有效保证金及法定保险/替代账户；状态为EXPIRED本身不足以认定例外成立。'),
        dict(clause='18.27.040(1),(2)',pdf_pages='7-8',topic='保证金及新额度',finding='一般承包商30000美元、专业15000美元；2024调整有续期过渡，仍须核对登记历史和适用范围。'),
        dict(clause='18.27.040(7),(8)',pdf_pages='9',topic='减损及替代担保',finding='无满足全额的未减损担保可能触发暂停；法条允许指定储蓄账户，单一保证金CSV不足以排除替代安排。'),
        dict(clause='18.27.050(1)',pdf_pages='10',topic='保险保障形式',finding='财产50000、单人伤害100000、多人伤害200000；单一insuranceamt不足以核验各项或保单形式。'),
        dict(clause='18.27.050(2),(3)',pdf_pages='10-11',topic='取消后果与财务责任替代',finding='届满/取消等有登记后果，但记录不全或已批准替代账户不能由本包推断不存在。'),
        dict(clause='18.27.090(14)',pdf_pages='14',topic='专业身份例外',finding='特定电气/管道专业身份且在认证范围内有例外；当前登记类型不能替代具体工作事实。'),
        dict(clause='18.27.010(17)',pdf_pages='4',topic='登记核验',finding='法条定义当前登记卡、数据库或向部门核验，下载快照不等于此后实时状态。'),
    ]
    return dict(summary=dict(general_rows=len(general),bond_rows=len(bonds),insurance_rows=len(insurance),
                             duplicate_license_keys=sum(v>1 for v in general_keys.values()),
                             active_current_both_gap=sum(r['active_current_both_gap'] for r in licenses),
                             active_renewal_gap=sum(r['active_renewal_gap'] for r in licenses),
                             policy_orphans=len(orphans),license_types=dict(Counter(r['contractorlicensetypecode'] for r in general)),insurance_legal_compliance_not_determined=True),
                licenses=licenses,policies=policies,orphans=orphans,identity_exceptions=exceptions,legal_matrix=legal_matrix,
                multiple_licenses_per_ubi=[dict(ubi=k,licenses=sorted(v)) for k,v in sorted(ubi_licenses.items()) if len(v)>1],
                current_gap=sorted([r for r in licenses if r['active_current_both_gap']],key=lambda r:r['license']),
                renewal_gap=sorted([r for r in licenses if r['active_renewal_gap']],key=lambda r:r['license']))


def da(inp):
    a = records(inp/'CAINC1__ALL_AREAS.csv', 'cp1252')
    b = records(inp/'CAINC4__ALL_AREAS.csv', 'cp1252')
    pop = records(inp/'co-est2024-alldata.csv', 'cp1252')
    with (inp/'national_county2020.txt').open(encoding='utf-8-sig') as stream:
        old = list(csv.DictReader(stream,delimiter='|'))
    old_codes = {r['STATEFP']+r['COUNTYFP'] for r in old}
    def index(rows):
        result = {}
        footnotes = []
        for r in rows:
            key = (r.get('GeoFIPS') or '').strip().strip('"').strip()
            line = r.get('LineCode') or ''
            if re.fullmatch(r'\d{5}', key) and line.isdigit():
                if (key,line) in result:
                    raise ValueError('Duplicate BEA key')
                result[key,line]=r
            else:
                footnotes.append(r['_row'])
        return result,footnotes
    ia,fa=index(a);ib,fb=index(b)
    cp={r['STATE'].zfill(2)+r['COUNTY'].zfill(3):r for r in pop if r['SUMLEV']=='050'}
    ga={key for key,line in ia if line=='1'}
    gb={key for key,line in ib if line=='10'}
    comparison=[]; exclusions=[]; checks=[]
    for code in sorted(cp):
        if code not in ga or code not in gb:
            exclusions.append(dict(fips=code,reason='missing BEA geography',census_row=cp[code]['_row']))
            continue
        row=dict(fips=code,name=ia[code,'1']['GeoName'],census_name=cp[code]['CTYNAME'],old_code_present=code in old_codes,
                 source_rows={'cainc1_income':ia[code,'1']['_row'],'cainc1_population':ia[code,'2']['_row'],
                              'cainc1_pc':ia[code,'3']['_row'],'census':cp[code]['_row'],
                              **{f'cainc4_{k}':ib[code,k]['_row'] for k in ['10','45','46','47']}})
        endpoints={}
        reasons=[]
        for year in [str(y) for y in range(2020,2025)]:
            pi=number(ia[code,'1'][year]); pi4=number(ib[code,'10'][year])
            p=number(ia[code,'2'][year]);c=number(cp[code]['POPESTIMATE'+year]);published=number(ia[code,'3'][year])
            parts={k:number(ib[code,k][year]) for k in ['45','46','47']}
            pc=pi*1000/p if pi is not None and p and p>0 else None
            pc_c=pi*1000/c if pi is not None and c and c>0 else None
            residual=pi4-sum(parts.values()) if pi4 is not None and all(v is not None for v in parts.values()) else None
            checks.append(dict(fips=code,year=year,income1=pi,income4=pi4,income_equal=pi==pi4,
                               bea_population=p,census_population=c,published_pc=published,recomputed_bea_pc=pc,
                               pc_rounding_residual=published-pc if published is not None and pc is not None else None,
                               component_residual=residual))
            if year in ['2020','2024']:
                endpoints[year]=dict(income=pi,income4=pi4,bea_pop=p,census_pop=c,published_pc=published,bea_pc=pc,census_pc=pc_c,parts=parts)
                if any(v is None or v<=0 for v in [pi,p,c,*parts.values()]):
                    reasons.append(year+':nonpositive_or_missing_endpoint')
                if pi!=pi4:
                    reasons.append(year+':income_tables_disagree')
        row['endpoints']=endpoints
        first,last=endpoints['2020'],endpoints['2024']
        row['population_difference']=last['bea_pop']-last['census_pop']
        row['population_difference_pct']=ratio(row['population_difference'],last['census_pop'])
        row['eligible']=not reasons
        if reasons:
            exclusions.append(dict(fips=code,reason=reasons))
        else:
            row.update(income_growth_pct=ratio(last['income']-first['income'],first['income']),
                       bea_pc_growth_pct=ratio(last['bea_pc']-first['bea_pc'],first['bea_pc']),
                       census_pc_growth_pct=ratio(last['census_pc']-first['census_pc'],first['census_pc']),
                       bea_population_growth_pct=ratio(last['bea_pop']-first['bea_pop'],first['bea_pop']),
                       census_population_growth_pct=ratio(last['census_pop']-first['census_pop'],first['census_pop']),
                       component_contribution_pp={k:ratio(last['parts'][k]-first['parts'][k],first['income']) for k in first['parts']},
                       log_income=math.log(float(last['income']/first['income'])),
                       log_bea_population=math.log(float(last['bea_pop']/first['bea_pop'])),
                       log_census_population=math.log(float(last['census_pop']/first['census_pop'])))
        comparison.append(row)
    eligible=[r for r in comparison if r['eligible']]
    ranks={basis:sorted(eligible,key=lambda r:(-r[basis+'_pc_growth_pct'],r['fips']))[:20] for basis in ['bea','census']}
    return dict(summary=dict(cainc1_rows=len(a),cainc4_rows=len(b),census_counties=len(cp),
                             directly_matched=len(comparison),eligible=len(eligible),
                             income_cross_table_mismatches=sum(not r['income_equal'] for r in checks),
                             nonzero_component_residuals=sum(r['component_residual'] not in (None,0) for r in checks),
                             footnotes={'cainc1':fa,'cainc4':fb}),
                coverage=dict(census_not_bea1=sorted(cp.keys()-ga),bea1_not_census=sorted(ga-cp.keys()),
                              census_not_bea4=sorted(cp.keys()-gb),bea4_not_census=sorted(gb-cp.keys()),
                              census_not_old_codes=sorted(cp.keys()-old_codes),old_codes_not_census=sorted(old_codes-cp.keys())),
                checks=checks,comparison=comparison,exclusions=exclusions,
                population_difference_top20=sorted(comparison,key=lambda r:(-abs(r['population_difference_pct']),r['fips']))[:20],
                rankings=ranks,rank_top20_overlap=sorted({r['fips'] for r in ranks['bea']} & {r['fips'] for r in ranks['census']}))


def ga(inp):
    sys.path.insert(0,str(inp/'data/tools'))
    from wesnoth.wmliterator3 import WmlIterator
    def parse(path):
        return WmlIterator(filename=str(path))
    def attrs(path, scope):
        out={}
        for row in parse(path):
            if [s.element for s in row.scopes]==scope and isinstance(row.element,str) and row.element.endswith('='):
                out[row.element[:-1]]=dict(value=row.text.split('=',1)[1].strip().strip('"'),line=row.lineno+1)
        return out
    faction_path=inp/'data/multiplayer/factions/loyalists-default.cfg'
    faction=attrs(faction_path,['[multiplayer_side]'])
    recruits=faction['recruit']['value'].split(',')
    leaders=faction['leader']['value'].split(',')
    units={}
    for path in sorted((inp/'data/core/units').rglob('*.cfg')):
        fields=attrs(path,['[unit_type]'])
        if not fields or 'id' not in fields:
            continue
        uid=fields['id']['value']
        if uid in units:
            raise ValueError('Duplicate unit ID '+uid)
        attacks=defaultdict(dict)
        overrides=defaultdict(dict)
        macros=[]
        for row in parse(path):
            scopes=[s.element for s in row.scopes]
            elements=row.element if isinstance(row.element,(list,tuple)) else [row.element]
            if scopes[:1]==['[unit_type]'] and ('[abilities]' in scopes or '[specials]' in scopes):
                for elem in elements:
                    if isinstance(elem,str) and elem.startswith(('{ABILITY_','{WEAPON_SPECIAL_')):
                        macros.append(dict(name=elem[1:],line=row.lineno+1))
            if len(scopes)==2 and scopes[0]=='[unit_type]' and isinstance(row.element,str) and row.element.endswith('='):
                value=row.text.split('=',1)[1].strip().strip('"')
                item=dict(value=value,line=row.lineno+1)
                if scopes[1]=='[attack]':
                    attacks[row.scopes[-1].lineno][row.element[:-1]]=item
                elif scopes[1] in ['[resistance]','[defense]','[movement_costs]']:
                    overrides[scopes[1][1:-1]][row.element[:-1]]=item
        units[uid]=dict(id=uid,path=str(path.relative_to(inp)),fields=fields,attacks=list(attacks.values()),
                        overrides=dict(overrides),macros=macros)
    reachable=set();edges=[];unresolved=[]
    def visit(uid,stack):
        if uid in stack:
            raise ValueError('Advancement cycle '+uid)
        if uid in reachable:
            return
        if uid not in units:
            unresolved.append(uid);return
        reachable.add(uid)
        targets=units[uid]['fields'].get('advances_to',{}).get('value','null')
        for target in targets.split(','):
            if target in ('null',''):
                continue
            edges.append((uid,target))
            visit(target,stack+[uid])
    for uid in recruits:
        visit(uid,[])
    edge_records=[]
    for source,target in edges:
        if target not in units:
            continue
        a,b=units[source],units[target]
        diffs={}
        for key in ['hitpoints','movement','cost']:
            va=number(a['fields'].get(key,{}).get('value'));vb=number(b['fields'].get(key,{}).get('value'))
            diffs[key]=vb-va if va is not None and vb is not None else None
        def attack_map(unit):
            return {tuple(attack.get(k,{}).get('value') for k in ['name','range','type']):attack for attack in unit['attacks']}
        ma,mb=attack_map(a),attack_map(b)
        changes=[]
        for key in sorted(ma.keys() | mb.keys()):
            values={}
            for field in ['damage','number']:
                va=number(ma.get(key,{}).get(field,{}).get('value'));vb=number(mb.get(key,{}).get(field,{}).get('value'))
                values[field]=dict(before=va,after=vb,delta=vb-va if va is not None and vb is not None else None)
            changes.append(dict(key=key,changes=values))
        edge_records.append(dict(source=source,target=target,deltas=diffs,attack_changes=changes,
                                 alignment_changed=a['fields']['alignment']['value']!=b['fields']['alignment']['value'],
                                 hp_up_movement_down=bool(diffs['hitpoints'] is not None and diffs['hitpoints']>0 and diffs['movement'] is not None and diffs['movement']<0)))
    paths=[]
    adj=defaultdict(list)
    for a,b in edges:
        adj[a].append(b)
    def expand(origin,uid,route,xp):
        paths.append(dict(origin=origin,target=uid,path=route,total_xp=xp))
        for nxt in adj[uid]:
            if nxt in route or nxt not in units:
                continue
            cost=number(units[uid]['fields'].get('experience',{}).get('value'))
            if cost is None:
                raise ValueError('Unresolved base XP '+uid)
            expand(origin,nxt,route+[nxt],xp+cost)
    for uid in recruits:
        expand(uid,uid,[uid],Decimal(0))
    movement_types={}
    current=None
    for row in parse(inp/'data/core/units.cfg'):
        scopes=[s.element for s in row.scopes]
        if scopes[:1]==['[units]']:
            scopes=scopes[1:]
        if row.element=='[movetype]':
            current={'properties':{},'defense':{},'movement_costs':{},'resistance':{}}
        if not current or not isinstance(row.element,str) or not row.element.endswith('='):
            continue
        value=row.text.split('=',1)[1].strip().strip('"')
        key=row.element[:-1]
        if scopes==['[movetype]']:
            current['properties'][key]=dict(value=value,line=row.lineno+1)
            if key=='name':
                movement_types[value]=current
        elif len(scopes)==2 and scopes[0]=='[movetype]' and scopes[1][1:-1] in current:
            current[scopes[1][1:-1]][key]=dict(value=value,line=row.lineno+1)
    terrain=[]
    for uid in sorted(reachable):
        unit=units[uid]
        mt=unit['fields'].get('movement_type',{}).get('value')
        base=movement_types.get(mt)
        if base is None:
            unresolved.append('movement_type:'+str(mt));continue
        values={}
        for kind,keys in [('defense',['flat','forest','village','shallow_water','deep_water']),
                          ('movement_costs',['flat','forest','village','shallow_water','deep_water']),
                          ('resistance',['blade','pierce','impact','fire','cold','arcane'])]:
            values[kind]={}
            for key in keys:
                inherited=base[kind].get(key)
                override=unit['overrides'].get(kind,{}).get(key)
                chosen=override or inherited
                values[kind][key]=dict(value=chosen['value'] if chosen else None,
                                      path=unit['path'] if override else 'data/core/units.cfg',
                                      line=chosen['line'] if chosen else None,unit_override=bool(override),
                                      limitation=None if chosen else 'No explicit value in movement_type or unit override; no engine-default value fabricated')
        terrain.append(dict(unit=uid,movement_type=mt,values=values))
    definitions={}
    for path in sorted((inp/'data/core/macros').glob('*.cfg')):
        lines=path.read_text().splitlines()
        for i,line in enumerate(lines):
            if line.startswith('#define '):
                name=line.split()[1]
                end=next((j+1 for j in range(i+1,len(lines)) if lines[j].startswith('#enddef')),len(lines))
                definitions[name]=dict(path=str(path.relative_to(inp)),line=i+1,end_line=end)
    abilities=[]
    for uid in sorted(reachable):
        for macro in units[uid]['macros']:
            name=macro['name']
            abilities.append(dict(unit=uid,macro=name,unit_path=units[uid]['path'],line=macro['line'],definition=definitions.get(name)))
    min_paths=[]
    by_target=defaultdict(list)
    for r in paths:
        by_target[r['origin'],r['target']].append(r)
    for key,rows in sorted(by_target.items()):
        best=min(r['total_xp'] for r in rows)
        min_paths.append(dict(origin=key[0],target=key[1],min_xp=best,paths=[r['path'] for r in rows if r['total_xp']==best],path_count=len(rows)))
    return dict(summary=dict(recruits=recruits,leaders=leaders,reachable_units=len(reachable),advancement_edges=len(edges),
                             paths=len(paths),unresolved=sorted(set(unresolved)),
                             hp_up_movement_down=sum(r['hp_up_movement_down'] for r in edge_records),
                             alignment_changed=sum(r['alignment_changed'] for r in edge_records)),
                units=[units[k] for k in sorted(reachable)],edges=edge_records,paths=paths,min_paths=min_paths,
                terrain=terrain,abilities=abilities)


def encode(value):
    if isinstance(value,Decimal):
        return float(value)
    raise TypeError(type(value).__name__)


def narrative(domain,result):
    if domain=='GO':
        s=result['summary']
        examples=result['top_ratios'][:3]
        return [('建议与已核实事实',[
            f'先按不同付款粒度分别补档核验，不报告去重后的真实总支出。两份快照有{s["payment_rows"]}条付款和{s["contract_version_rows"]}条合同版本记录。',
            f'合同{", ".join(s["grain_overlap"])}同时存在两种付款粒度，需要调取明细和年度汇总定义确认是否重叠。',
            f'直接选全快照最新记录会改变{s["future_changed_contracts"]}个合同的年末可见视图，其中{s["paid_future_changed_contracts"]}个涉及本付款快照。']),
            ('真实追查案例',[
                f'合同{r["contract"]}：年度付款{r["amount"]}美元，授予额{r["award"]}美元，比率{r["paid_award_pct"]:.4f}%；付款行{r["payment_rows"]}，选中合同原始行{r["selected_contract_rows"]}，批准{r["approval_date"]}，修订{r["revision"]}。需取得完整合同、累计执行和付款归属，当前比率不是违法证明。'
                for r in examples]),
            ('后续核验与限制',['优先联系原始数据发布方核对混合粒度，再核对未匹配及版本冲突记录。合同表始于2019年且不是完整采购档案，付款没有发票日期；供应商写法差异不能作为强配实体的依据。'])]
    if domain=='FIN':
        s=result['summary']
        return [('结论与分析范围',[
            f'按真实报告集合开展尽调，而不是先按当前ACTIVE筛样本。面板有{s["observations"]}条银行季度记录、{s["banks"]}家银行、{len(s["quarters"])}季；完整窗口固定样本为{s["fixed_banks"]}家。',
            '资产质量代理仅为90天以上逾期加非应计资产/总资产；金额原单位千美元，代理不是监管不良贷款率。固定样本、动态样本和当前存续筛选必须分别报告。']),
            ('前三家追查',[
                f'{r["name"]}（CERT {r["cert"]}）：代理同比变化{r["yoy_pp"]:.6f}个百分点，末期{r["quality_pct"]:.6f}%；两期原始行{r["source_rows"]}。先取得逾期/非应计资产明细、抵押品、评级迁徙和资本计算资料，核对该变化是否来自个别敞口、处置或分母变化。'
                for r in result['top10'][:3]]),
            ('手册依据与不足',[
                'FDIC_RMS_Manual_3.2_Loans.pdf，PDF第1-3页的风险管理/贷款政策讨论及第11-12页concentrations相关内容：组合金额与集中度应与风险管理、具体贷款质量一起审查，字段比率不足以完成个案评级。',
                'OCC_Rating_Credit_Risk.pdf，PDF第9页及第24页涉及定量因素的讨论：本表缺少借款人现金流、抵押品及完整信用判断，不能用单一代理给出监管评级。',
                '机构表是较晚快照，ACTIVE并不重建历史每季资格。本分析是快照回顾，未证明2025年时点可得性。'])]
    if domain=='LAW':
        s=result['summary']
        cases=result['current_gap'][:3]+result['renewal_gap'][:2]
        sections=[('建议与结论',[
            f'将当前ACTIVE但两类保障均无可确认日期记录的{s["active_current_both_gap"]}个执照交人工核验，将未来30天已记录保障存在后继缺口的{s["active_renewal_gap"]}个执照交续期核对。它们是证据待补名单，不是违法名单。',
            f'登记类型分布为{json.dumps(s["license_types"],ensure_ascii=False)}。不能把建筑登记法的门槛机械套给其他专业类型；其实际工作及豁免适用仍需核实。']),
            ('法律依据',[f'RCW {r["clause"]}，RCW_18.27.pdf第{r["pdf_pages"]}页：{r["finding"]}' for r in result['legal_matrix']]),
            ('真实记录案例',[f'{r["business"]}，执照{r["license"]}、UBI {r["ubi"]}，登记原始行{r["source_row"]}，表列状态{r["reported_status"]}；保证金行{r["bond_rows"]}、保险行{r["insurance_rows"]}。核对逐保障记录表的生效、取消与后继日期，不能用当前状态重写历史记录。' for r in cases]),
            ('事实和法律限制',['所有日期结论按题面筛查日界线计算，并非原始保单小时级法律终止裁定。单一保险金额不证明法定分项、无除外责任或覆盖范围，未见记录不能排除其他已批准财务责任安排。未取得合同、实际工作或通知事实，不认定终局法律责任。'])]
        for title,items,request in [('核验函草稿一',result['current_gap'],'请提供评估日有效的保证金、保险或部门批准的替代财务责任资料及登记证明。当前快照未能确认两类日期支持，尚未认定贵企业未投保或违法。'),('核验函草稿二',result['renewal_gap'],'请核对未来30天内保障终止日期与后继生效资料，提供续期或替代保障证明，并说明期间是否存在已覆盖但尚未反映在下载记录中的安排。')]:
            if items:
                r=items[0]
                sections.append((title,[f'致{r["business"]}资料核验联系人（未发送）：就公开执照{r["license"]}（登记原始行{r["source_row"]}）开展资料核对。{request} 此函不是行政处理或违法通知，资料补齐后再由专业人员核验适用范围。']))
        return sections
    if domain=='DA':
        s=result['summary'];diff=result['population_difference_top20'][:3]
        return [('发布结论',[
            f'发布范围限可直接按代码比较的地区；Census目标共{s["census_counties"]}县，实际直接匹配{s["directly_matched"]}县，满足增长贡献端点条件{s["eligible"]}县。其余保持在覆盖/排除记录，不靠名称补造映射。',
            f'两份排名前20交集为{len(result["rank_top20_overlap"])}个FIPS，应分别展示，不隐去人口口径对结果的影响。']),
            ('人口差异样例',[f'{r["fips"]} {r["name"]}：2024 BEA减Census人口差{r["population_difference"]}，相对差{r["population_difference_pct"]:.6f}%；两端数值和源行见JSON endpoints/source_rows。差异不自动证明任何一份公开原件错误。' for r in diff]),
            ('核验及使用限制',[
                f'2020—2024直接匹配记录中，跨表总收入不一致有{s["income_cross_table_mismatches"]}项，三分项非零加总残差有{s["nonzero_component_residuals"]}项。没有发现就报告0，不为增加难度捏造冲突。',
                'BEA金额为千美元、人均为美元，重算前必须乘1000。贡献为分项收入增量/基期总收入的百分点，不是各项自身增长率。对数分解是收入变化减人口变化；负值、抑制值及不可比地域不填0。',
                '现有县代码与历史2020代码并不等同；本包没有官方合并地域交叉表，不能自行拆分。所有增长均为名义值，不等于通胀调整后生活水平改善。'])]
    special=[r for r in result['edges'] if r['hp_up_movement_down'] or r['alignment_changed']]
    return [('审查结论',[
        f'忠诚派的{len(result["summary"]["recruits"])}个真实招募起点可达{result["summary"]["reachable_units"]}个单位、{result["summary"]["advancement_edges"]}条升级边。招募、领袖、可达目标分别列示。',
        '将静态升级目录用于确定游戏内验证问题，不把面板武器数值当完整对战胜率。路径XP是非终点基础experience之和，不包含实际击杀收益、traits、经验倍率或AMLA。']),
        ('存在权衡的真实升级边',[f'{r["source"]} → {r["target"]}：HP变化{r["deltas"]["hitpoints"]}、movement变化{r["deltas"]["movement"]}、alignment改变={r["alignment_changed"]}。源单位和目标单位的真实属性/武器/覆盖参见工作簿。' for r in special]),
        ('能力依赖与建议实测',[
            'Horseman与骑士路线的charge须结合攻击/反击条件检验；不能把面板值无条件乘二作为所有场景的净收益。建议比较实际冲锋时对方反击距离和生存风险。',
            'Spearman的firststrike影响交互顺序，不能当作每次攻击都固定额外伤害；建议在可能提前击杀的真实条件下用引擎核验顺序与收益。',
            'Mage路线的magical影响命中规则，不能以damage×number/cost替代地形和抗性相关评价；建议比较法师各分支在高防御地形、不同伤害抗性条件下的表现。',
            '三个验证问题均未在本题实际运行游戏。原始防御为被命中比例，抗性配置表示承受伤害比例，负防御还涉及上限；具体解释应回查src/movetype.cpp，不填补原件中未显式提供的地形默认值。'])]


def write_office(out,domain,result):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from docx import Document
    sheet_map={
        'GO':[('口径与守恒','conservation'),('版本对照','versions_changed'),('未匹配追查','top_unmatched'),('高比率追查','top_ratios'),('例外','exceptions')],
        'FIN':[('覆盖核对','coverage'),('季度面板','panel'),('动态与固定样本','aggregates'),('存续偏差','endpoints'),('尽调名单','top10')],
        'LAW':[('法律矩阵','legal_matrix'),('逐执照','licenses'),('逐保障记录','policies'),('待核验与续期','renewal_gap'),('当前双缺口','current_gap'),('数据例外','identity_exceptions')],
        'DA':[('覆盖与例外','exclusions'),('跨表核验','checks'),('人口差异','population_difference_top20'),('双口径排名','rankings'),('增长贡献','comparison')],
        'GA':[('单位目录','units'),('升级边','edges'),('路径XP','min_paths'),('地形抗性','terrain'),('能力依赖与覆盖','abilities')],
    }
    def flatten(value,prefix=''):
        answer={}
        for key,item in value.items():
            name=prefix+str(key)
            if isinstance(item,dict):
                answer.update(flatten(item,name+'.'))
            elif isinstance(item,list):
                answer[name]=json.dumps(item,ensure_ascii=False,default=encode)
            else:
                answer[name]=float(item) if isinstance(item,Decimal) else item
        return answer
    book=Workbook();book.remove(book.active)
    for title,key in sheet_map[domain]:
        rows=result[key]
        if isinstance(rows,dict):
            rows=[dict(basis=basis,**r) for basis,items in rows.items() for r in items]
        values=[flatten(row) for row in rows]
        headers=list(dict.fromkeys(k for row in values for k in row)) or ['说明']
        sheet=book.create_sheet(title);sheet.append(headers)
        for row in values:
            sheet.append([row.get(h) for h in headers])
        sheet.freeze_panes='A2';sheet.auto_filter.ref=sheet.dimensions
        for index,h in enumerate(headers,1):
            sheet.column_dimensions[get_column_letter(index)].width=min(48,max(15,len(h)+2))
        for cell in sheet[1]:
            cell.font=Font(bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='275B49')
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment=Alignment(vertical='top',wrap_text=True)
                if isinstance(cell.value,float):cell.number_format='0.0000'
    book.save(out/'audit.xlsx')
    doc=Document();doc.add_heading(domain+' 真实来源参考意见',0)
    for heading,paragraphs in narrative(domain,result):
        doc.add_heading(heading,1)
        for paragraph in paragraphs:doc.add_paragraph(paragraph)
    doc.add_heading('参考稿状态',1)
    doc.add_paragraph('本稿与配套代码为作者参考解，待独立领域审核及正式判分；不声称Gold已验收。输入为保留原始字节的公开快照，来源记录在题包provenance中。')
    doc.save(out/'memo.docx')


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--domain',choices=['GO','FIN','LAW','DA','GA'])
    parser.add_argument('--input',type=Path,required=True);parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args()
    domain=args.domain
    if domain is None:
        domain=next(d for d,name in [('GO','Chicago_Payments_2024.csv'),('FIN','fdic_financials_VT_NH_ME_2022_2025.csv'),
                                    ('LAW','LNI_Contractor_License_General_BR.csv'),('DA','CAINC1__ALL_AREAS.csv'),
                                    ('GA','data/multiplayer/factions/loyalists-default.cfg')] if (args.input/name).is_file())
    result={'GO':go,'FIN':fin,'LAW':law,'DA':da,'GA':ga}[domain](args.input)
    args.output.mkdir(parents=True,exist_ok=True)
    (args.output/'results.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,default=encode))
    write_office(args.output,domain,result)
    print(json.dumps(result['summary'],ensure_ascii=False,default=encode))
