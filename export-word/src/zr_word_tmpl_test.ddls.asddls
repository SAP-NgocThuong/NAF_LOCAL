@AccessControl.authorizationCheck: #NOT_REQUIRED
@EndUserText.label: 'Root view Word template test'
define root view entity ZR_WORD_TMPL_TEST
  as select from ztb_wdtpl_test
{
  key tmpl_uuid             as TmplUuid,
      tmpl_id               as TmplId,
      description           as Description,

      @Semantics.largeObject: { mimeType: 'TmplMime',
                                fileName: 'TmplName',
                                acceptableMimeTypes: [ 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                                                       'image/png', 'image/jpeg' ],
                                contentDispositionPreference: #ATTACHMENT }
      tmpl_content          as TmplContent,
      @Semantics.mimeType: true
      tmpl_mime             as TmplMime,
      tmpl_name             as TmplName,

      @Semantics.user.createdBy: true
      created_by            as CreatedBy,
      @Semantics.systemDateTime.createdAt: true
      created_at            as CreatedAt,
      @Semantics.user.localInstanceLastChangedBy: true
      local_last_changed_by as LocalLastChangedBy,
      @Semantics.systemDateTime.localInstanceLastChangedAt: true
      local_last_changed_at as LocalLastChangedAt,
      @Semantics.systemDateTime.lastChangedAt: true
      last_changed_at       as LastChangedAt
}
