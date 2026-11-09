-- Stage 1 -> amir_lab.tam_rollup : contact-database rollup per domain (size + coverage).
-- n=contacts, nv=valid/catch-all-valid emails, nx=exec contacts, ns=sales/ISO/BD contacts, cn=company name,
-- sb=size bucket, cs=max company size, rb=revenue bucket, ar=max annual revenue, il=has lending industry tag.
WITH '(small business|business) (lend|loan|financ|fund|capital)|merchant (cash|fund)|cash advance|\\bmca\\b|working capital|revenue.based (financ|fund)|equipment (financ|leas)|invoice (financ|factor)|\\bfactoring\\b|alternative (lend|financ)|commercial (lend|financ)|sba (lend|loan|7a)|loan brokerage|loan broker|funding marketplace|loan marketplace|business credit|lines? of credit|asset.based lend|funding solutions' AS rx_lend,
     'mortgage|real estate|home loan|student|auto loan|car loan|personal loan|payday|pawn|title loan|agricultur|ymca|notary|loan signing' AS rx_excl
SELECT d AS domain, count() AS n,
       countIf(verification_status IN ('valid', 'catch_all_valid')) AS nv,
       countIf(match(lower(concat(ifNull(seniority, ''), ' ', ifNull(job_title, ''))), 'owner|founder|c_suite|c-suite|c-team|\\bceo\\b|president|chief|principal|managing (member|director|partner)')) AS nx,
       countIf(match(lower(ifNull(job_title, '')), '\\biso\\b|partner relations|broker relations|channel|business development|sales')) AS ns,
       topK(1)(company_name)[1] AS cn, topK(1)(company_size_bucket)[1] AS sb, max(company_size) AS cs,
       topK(1)(annual_revenue_bucket)[1] AS rb, max(annual_revenue) AS ar,
       max(match(lower(ifNull(specific_industry, '')), rx_lend) AND NOT match(lower(ifNull(specific_industry, '')), rx_excl)) AS il
FROM (SELECT *, lower(splitByChar('@', email)[2]) AS d FROM mirror.leads_current WHERE email LIKE '%@%')
WHERE d IN (SELECT domain FROM {universe})
GROUP BY d
