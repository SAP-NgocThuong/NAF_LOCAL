@AccessControl.authorizationCheck: #NOT_REQUIRED
@EndUserText.label: 'Projection view Word template test'
@Metadata.allowExtensions: true
define root view entity ZC_WORD_TMPL_TEST
  provider contract transactional_query
  as projection on ZR_WORD_TMPL_TEST
{
  key TmplUuid,
      TmplId,
      Description,
      @Semantics.largeObject: { mimeType: 'TmplMime',
                                fileName: 'TmplName',
                                acceptableMimeTypes: [ 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ],
                                contentDispositionPreference: #ATTACHMENT }
      TmplContent,
      @Semantics.mimeType: true
      TmplMime,
      TmplName,
      CreatedBy,
      CreatedAt,
      LocalLastChangedBy,
      LocalLastChangedAt,
      LastChangedAt
}
