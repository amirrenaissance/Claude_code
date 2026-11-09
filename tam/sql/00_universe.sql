-- Stage 0 -> amir_lab.tam_universe : every candidate Business Lending domain.
-- Sources: lending vertical tag, website sub-industry/one-liner, contact-industry tag, Attio keyword hits + Partners ({extra_domains}).
-- Leads-only domains are noisy; build_tam.py keeps them only when the name/domain shows lending evidence.
WITH '(small business|business) (lend|loan|financ|fund|capital)|merchant (cash|fund)|cash advance|\\bmca\\b|working capital|revenue.based (financ|fund)|equipment (financ|leas)|invoice (financ|factor)|\\bfactoring\\b|alternative (lend|financ)|commercial (lend|financ)|sba (lend|loan|7a)|loan brokerage|loan broker|funding marketplace|loan marketplace|business credit|lines? of credit|asset.based lend|funding solutions' AS rx_lend,
     'mortgage|real estate|home loan|student|auto loan|car loan|personal loan|payday|pawn|title loan|agricultur|ymca|notary|loan signing' AS rx_excl
SELECT d AS domain, max(s = 'vertical') AS in_vertical, max(s = 'site') AS in_site, max(s = 'leads') AS in_leads, max(s = 'attio') AS in_attio
FROM (
    SELECT lower(domain) AS d, 'vertical' AS s FROM mirror.domain_vertical WHERE vertical = 'mca_alt_lending'
    UNION ALL
    SELECT lower(domain), 'site' FROM mirror.company_site_profile
    WHERE match(lower(concat(ifNull(sub_industry, ''), ' ', ifNull(one_liner, ''))), rx_lend) AND NOT match(lower(ifNull(sub_industry, '')), rx_excl)
    UNION ALL
    SELECT DISTINCT lower(splitByChar('@', email)[2]), 'leads' FROM mirror.leads_current
    WHERE email LIKE '%@%' AND match(lower(ifNull(specific_industry, '')), rx_lend) AND NOT match(lower(ifNull(specific_industry, '')), rx_excl)
    UNION ALL
    SELECT arrayJoin({extra_domains}), 'attio'
)
WHERE d != '' AND position(d, '.') > 0 AND d NOT IN (SELECT lower(domain) FROM mirror.personal_email_domain)
GROUP BY d
