@EndUserText.label : 'Word template test'
@AbapCatalog.enhancement.category : #NOT_EXTENSIBLE
@AbapCatalog.tableCategory : #TRANSPARENT
@AbapCatalog.deliveryClass : #A
@AbapCatalog.dataMaintenance : #RESTRICTED
define table ztb_wdtpl_test {

  key client            : abap.clnt not null;
  key tmpl_uuid         : sysuuid_x16 not null;
  tmpl_id               : abap.char(20);
  description           : abap.char(80);
  tmpl_content          : abap.rawstring(0);
  tmpl_name             : abap.char(128);
  tmpl_mime             : abap.char(128);
  created_by            : abp_creation_user;
  created_at            : abp_creation_tstmpl;
  local_last_changed_by : abp_locinst_lastchange_user;
  local_last_changed_at : abp_locinst_lastchange_tstmpl;
  last_changed_at       : abp_lastchange_tstmpl;

}
