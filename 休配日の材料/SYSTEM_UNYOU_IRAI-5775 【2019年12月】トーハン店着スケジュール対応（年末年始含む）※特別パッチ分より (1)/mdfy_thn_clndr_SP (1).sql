--SEJ
UPDATE t_deliver_arng_dy_clndr_mst
SET    DELIVER_ARNG_DY_TO  = TO_CHAR(TO_DATE(DELIVER_ARNG_DY_TO) + 1,'YYYYMMDD')
      ,DELIVER_ARNG_DY_TO_FKSU = TO_CHAR(TO_DATE(DELIVER_ARNG_DY_TO_FKSU) + 1,'YYYYMMDD')
      ,UNYO_RIYU = '特別パッチ お届け予定日TOに＋１加算'
      ,UNYOTMP = SYSDATE
      ,UNYOUSR_ID = 'UNYO'
WHERE  TRUNC(YMD) > TRUNC(SYSDATE)
AND    NVL(UNYO_RIYU,'.') NOT LIKE '%特別パッチ%'
AND    accpt_kbn = '01'
AND    accpt_tnpo_jigyo_cmpny_cd = '011'  --SEJ
AND    shipment_arng_dy_from IS NOT NULL
AND    shipment_arng_dy_to IS NOT NULL
AND    DELIVER_ARNG_DY_FROM <> DELIVER_ARNG_DY_TO
AND(
	   (SUBSTR(snd_region_cd,1,2) IN ('47')
        AND  ((tohan_leadtm_from =  0 AND to_date(shipment_arng_dy_from) -1 = '2019/12/30')
        or (tohan_leadtm_from <> 0 AND to_date(shipment_arng_dy_from) = '2019/12/30')))
    or
	   (SUBSTR(snd_region_cd,1,2) IN ('31','32','33','34','35','40','41','42','43','44','45','46')
        AND  ((tohan_leadtm_from =  0 AND to_date(shipment_arng_dy_from) -1 = '2020/01/01')
        or (tohan_leadtm_from <> 0 AND to_date(shipment_arng_dy_from) = '2020/01/01')))
    or
	   (SUBSTR(snd_region_cd,1,2) IN ('01','02','36','37','38','39')
        AND  ((tohan_leadtm_from =  0 AND to_date(shipment_arng_dy_from) -1 = '2020/01/02')
        or (tohan_leadtm_from <> 0 AND to_date(shipment_arng_dy_from) = '2020/01/02')))
    or
	   (SUBSTR(snd_region_cd,1,2) IN ('03','04','05','06','07','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30')
        AND  ((tohan_leadtm_from =  0 AND to_date(shipment_arng_dy_from) -1 = '2020/01/03')
        or (tohan_leadtm_from <> 0 AND to_date(shipment_arng_dy_from) = '2020/01/03')))
    )
;


--SEJ以外
UPDATE t_deliver_arng_dy_clndr_mst
SET    DELIVER_ARNG_DY_TO  = TO_CHAR(TO_DATE(DELIVER_ARNG_DY_TO) + 1,'YYYYMMDD')
      ,DELIVER_ARNG_DY_TO_FKSU = TO_CHAR(TO_DATE(DELIVER_ARNG_DY_TO_FKSU) + 1,'YYYYMMDD')
      ,UNYO_RIYU = '特別パッチ お届け予定日TOに＋１加算'
      ,UNYOTMP = SYSDATE
      ,UNYOUSR_ID = 'UNYO'
WHERE  TRUNC(YMD) > TRUNC(SYSDATE)
AND    NVL(UNYO_RIYU,'.') NOT LIKE '%特別パッチ%'
AND    accpt_kbn = '01'
AND    accpt_tnpo_jigyo_cmpny_cd <> '011'
AND    shipment_arng_dy_from IS NOT NULL
AND    shipment_arng_dy_to IS NOT NULL
AND    DELIVER_ARNG_DY_FROM <> DELIVER_ARNG_DY_TO
AND(    
       (SUBSTR(snd_region_cd,1,2) IN ('47')
        AND  ((tohan_leadtm_from =  0 AND to_date(shipment_arng_dy_from) -1 = '2019/12/30')
        OR (tohan_leadtm_from <> 0 AND to_date(shipment_arng_dy_from) = '2019/12/30')))
    or
       (SUBSTR(snd_region_cd,1,2) IN ('31','32','33','34','35','40','41','42','43','44','45','46')
        AND  ((tohan_leadtm_from =  0 AND to_date(shipment_arng_dy_from) -1 = '2020/01/01')
        OR (tohan_leadtm_from <> 0 AND to_date(shipment_arng_dy_from) = '2020/01/01')))
    or
       (SUBSTR(snd_region_cd,1,2) IN ('01','02','36','37','38','39')
        AND  ((tohan_leadtm_from =  0 AND to_date(shipment_arng_dy_from) -1 = '2020/01/02')
        OR (tohan_leadtm_from <> 0 AND to_date(shipment_arng_dy_from) = '2020/01/02')))
    or
       (SUBSTR(snd_region_cd,1,2) IN ('03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30')
        AND  ((tohan_leadtm_from =  0 AND to_date(shipment_arng_dy_from) -1 = '2020/01/03')
        OR (tohan_leadtm_from <> 0 AND to_date(shipment_arng_dy_from) = '2020/01/03')))
   )
;

COMMIT;

EXIT;


