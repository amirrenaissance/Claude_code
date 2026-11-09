-- Stage 2 -> amir_lab.tam_site : website profile per domain (classification text, staff, founded, HQ).
SELECT lower(domain) AS d, any(company_name) AS cn, any(sub_industry) AS si, any(substring(ifNull(one_liner, ''), 1, 140)) AS ol,
       any(customer_type) AS ct, any(employee_count_stated) AS es, any(year_founded) AS yf, any(hq_state) AS st,
       any(is_hiring) AS hi, any(length(leadership)) AS nl
FROM mirror.company_site_profile
WHERE lower(domain) IN (SELECT domain FROM {universe})
GROUP BY d
