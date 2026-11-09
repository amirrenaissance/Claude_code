-- Stage 3 -> amir_lab.tam_verify : free verification signals per domain.
-- pv=primary vertical, ss=website crawl status, mx=mail server status, age=domain age (months), dead=on defunct list.
WITH U AS (SELECT domain AS d FROM {universe})
SELECT U.d, any(v.vertical) AS pv, any(c.s) AS ss, any(m.s) AS mx, any(t.a) AS age, max(df.x) AS dead
FROM U
LEFT JOIN (SELECT lower(domain) AS d, vertical FROM mirror.domain_vertical_primary WHERE lower(domain) IN U) v ON v.d = U.d
LEFT JOIN (SELECT lower(domain) AS d, argMax(site_status, checked_on) AS s FROM mirror.site_check_domains WHERE lower(domain) IN U GROUP BY d) c ON c.d = U.d
LEFT JOIN (SELECT lower(email_domain) AS d, argMax(mx_status, checked_at) AS s FROM mirror.email_domain_mx WHERE lower(email_domain) IN U GROUP BY d) m ON m.d = U.d
LEFT JOIN (SELECT lower(domain) AS d, any(domain_age_months) AS a FROM mirror.tib_domain_mv WHERE lower(domain) IN U GROUP BY d) t ON t.d = U.d
LEFT JOIN (SELECT lower(domain) AS d, 1 AS x FROM mirror.defunct_domain WHERE lower(domain) IN U) df ON df.d = U.d
GROUP BY U.d
