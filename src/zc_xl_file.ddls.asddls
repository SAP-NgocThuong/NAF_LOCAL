@AccessControl.authorizationCheck: #NOT_REQUIRED
@EndUserText.label: 'File store (Excel, Word, image) proj'
@Metadata.allowExtensions: true
define root view entity ZC_XL_FILE
  provider contract transactional_query
  as projection on ZR_XL_FILE
{
  key FileUuid,
      FileId,
      Version,
      Description,
      @Semantics.largeObject: { mimeType: 'MimeType',
                                fileName: 'FileName',
                                acceptableMimeTypes: [ 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                                       'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                                                       'image/png', 'image/jpeg' ],
                                contentDispositionPreference: #ATTACHMENT }
      FileContent,
      @Semantics.mimeType: true
      MimeType,
      FileName,
      CreatedBy,
      CreatedAt,
      LocalLastChangedBy,
      LocalLastChangedAt,
      LastChangedAt
}
