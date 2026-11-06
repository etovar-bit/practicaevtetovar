/*
Título: Columnas faltantes de verf.xlsx
Contexto: Completa la hoja verf (24 guías) con los datos que el Excel no trae,
          más el ID para cruzar cada fila. La razón social sale como en plataforma.

Qué hace: Por cada par (id_cliente, guía) del archivo verf.xlsx devuelve:
          - ID
          - Tipo de cliente
          - Razón social (persona física: nombre; persona moral: company_name)
          - Fecha de alta
          - Guía
          - Peso cliente (kg) y peso estafeta (kg)
          - Largo, ancho y alto declarados por el cliente (cm)
          - Envíos históricos, envíos con sobrepeso y monto histórico de sobrepeso
          - Producto declarado
          - Web / integración
          - Dirección de origen

Tablas:
- client_account
- transactions_account

Razón social:
- person_type_id = 1 (física): first_name, middle_name, last_name
- person_type_id = 2 (moral): company_name
- otro: nombre si existe; si no, company_name

Cómo ajustarlo:
- Cambia los pares en el subquery `pares` (id_cliente, guia).
- Repite los mismos id en el IN de `hist` y en los dos FIELD del ORDER BY
  para conservar el orden del Excel.
*/

SELECT
    pares.id_cliente AS `ID`,
    ca.client_type_id AS `Tipo de cliente`,
    CASE
        WHEN ca.person_type_id = 1 THEN
            CONCAT_WS(' ', ca.first_name, ca.middle_name, ca.last_name)
        WHEN ca.person_type_id = 2 THEN
            ca.company_name
        ELSE
            COALESCE(
                NULLIF(TRIM(CONCAT_WS(' ', ca.first_name, ca.middle_name, ca.last_name)), ''),
                ca.company_name
            )
    END AS `Razón social`,
    ca.created_at AS `Fecha de alta`,
    pares.guia AS `Guía`,
    ta.details_transactions->>'$.order.overweight_data.overweight.et_weight' AS `Peso cliente (kg)`,
    ta.details_transactions->>'$.order.overweight_data.overweight.provide_weight.weight' AS `Peso estafeta (kg)`,
    ta.details_transactions->>'$.order.package.length' AS `Largo cliente (cm)`,
    ta.details_transactions->>'$.order.package.width' AS `Ancho cliente (cm)`,
    ta.details_transactions->>'$.order.package.height' AS `Alto cliente (cm)`,
    IFNULL(hist.envios_historicos, 0) AS `Envíos históricos`,
    IFNULL(hist.envios_con_sobrepeso, 0) AS `Envíos con sobrepeso`,
    IFNULL(hist.monto_total_sobrepeso_historico, 0) AS `Monto Total Sobrepeso histórico($)`,
    ta.details_transactions->>'$.order.package.package_content' AS `Producto declarado`,
    ca.site AS `Web / integración`,
    CONCAT_WS(' ',
        NULLIF(TRIM(ta.details_transactions->>'$.order.address.origin.street'), ''),
        NULLIF(TRIM(ta.details_transactions->>'$.order.address.origin.ext_number'), ''),
        NULLIF(TRIM(ta.details_transactions->>'$.order.address.origin.int_number'), ''),
        NULLIF(TRIM(ta.details_transactions->>'$.order.address.origin.suburb'), ''),
        NULLIF(TRIM(ta.details_transactions->>'$.order.address.origin.state'), ''),
        NULLIF(TRIM(ta.details_transactions->>'$.order.address.origin.country_code'), '')
    ) AS `DIRECCION ORIGEN`
FROM (
    SELECT 12115 AS id_cliente, '1055892120610706592931' AS guia UNION ALL
    SELECT 49549, '1015892120610605242767' UNION ALL
    SELECT 83473, '1015892120610605283192' UNION ALL
    SELECT 61140, '8055892120610706419602' UNION ALL
    SELECT 70547, '2055892120610706472377' UNION ALL
    SELECT 52679, '2055892120610706570160' UNION ALL
    SELECT 65080, '3055892120610706433658' UNION ALL
    SELECT 37970, '8055892120610706595597' UNION ALL
    SELECT 99417, '3055892120610706590974' UNION ALL
    SELECT 94048, '4055892120610706446193' UNION ALL
    SELECT 58489, '6015892120610605282874' UNION ALL
    SELECT 92011, '1055892120610706595685' UNION ALL
    SELECT 63755, '1055892120610706597849' UNION ALL
    SELECT 96357, '8055892120610706597553' UNION ALL
    SELECT 59139, '1055892120610706594096' UNION ALL
    SELECT 87319, '4015892120610605280598' UNION ALL
    SELECT 86293, '6055892120610706589219' UNION ALL
    SELECT 75354, '0055892120610706595935' UNION ALL
    SELECT 63227, '4055892120610706547819' UNION ALL
    SELECT 62806, '6055892120610706548322' UNION ALL
    SELECT 90724, '6055892120610706597950' UNION ALL
    SELECT 99301, '6055892120610706596885' UNION ALL
    SELECT 94662, '1055892120610706592370' UNION ALL
    SELECT 94872, '9055892120610706582858'
) pares
LEFT JOIN client_account ca
       ON ca.id = pares.id_cliente
LEFT JOIN transactions_account ta
       ON ta.id = (
            SELECT t2.id
            FROM transactions_account t2
            WHERE t2.client_account_id = pares.id_cliente
              AND t2.provider_tracking_id = pares.guia
            ORDER BY
                (t2.details_transactions->>'$.order.overweight_data.overweight.provide_weight.weight' IS NOT NULL) DESC,
                t2.id DESC
            LIMIT 1
       )
LEFT JOIN (
    SELECT
        t.client_account_id,
        COUNT(*) AS envios_historicos,
        SUM(
            CASE
                WHEN JSON_EXTRACT(t.details_transactions, '$.order.overweight_data.overweight') IS NOT NULL
                THEN 1 ELSE 0
            END
        ) AS envios_con_sobrepeso,
        SUM(
            CASE
                WHEN JSON_EXTRACT(t.details_transactions, '$.order.overweight_data.overweight.charge.amount') IS NOT NULL
                THEN CAST(
                    NULLIF(t.details_transactions->>'$.order.overweight_data.overweight.charge.amount', '')
                    AS DECIMAL(12,2)
                )
                ELSE 0
            END
        ) AS monto_total_sobrepeso_historico
    FROM transactions_account t
    WHERE t.client_account_id IN (
        12115, 49549, 83473, 61140, 70547, 52679, 65080, 37970, 99417, 94048,
        58489, 92011, 63755, 96357, 59139, 87319, 86293, 75354, 63227, 62806,
        90724, 99301, 94662, 94872
    )
      AND t.status_id IN (41, 42, 43, 44, 59, 68, 69, 80, 84, 85, 86, 87, 88)
    GROUP BY t.client_account_id
) hist
       ON hist.client_account_id = pares.id_cliente
ORDER BY FIELD(pares.id_cliente,
        12115, 49549, 83473, 61140, 70547, 52679, 65080, 37970, 99417, 94048,
        58489, 92011, 63755, 96357, 59139, 87319, 86293, 75354, 63227, 62806,
        90724, 99301, 94662, 94872),
         FIELD(pares.guia,
        '1055892120610706592931', '1015892120610605242767', '1015892120610605283192',
        '8055892120610706419602', '2055892120610706472377', '2055892120610706570160',
        '3055892120610706433658', '8055892120610706595597', '3055892120610706590974',
        '4055892120610706446193', '6015892120610605282874', '1055892120610706595685',
        '1055892120610706597849', '8055892120610706597553', '1055892120610706594096',
        '4015892120610605280598', '6055892120610706589219', '0055892120610706595935',
        '4055892120610706547819', '6055892120610706548322', '6055892120610706597950',
        '6055892120610706596885', '1055892120610706592370', '9055892120610706582858');
