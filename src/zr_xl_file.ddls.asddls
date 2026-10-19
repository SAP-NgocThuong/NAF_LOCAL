@AccessControl.authorizationCheck: #NOT_REQUIRED
@EndUserText.label: 'File store (Excel, Word, image)'
define root view entity ZR_XL_FILE
  as select from zxl_file
{
  key file_uuid             as FileUuid,
      file_id               as FileId,
      version               as Version,
      description           as Description,

      @Semantics.largeObject: { mimeType: 'MimeType',
                                fileName: 'FileName',
                                acceptableMimeTypes: [ 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                                       'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                                                       'image/png', 'image/jpeg' ],
                                contentDispositionPreference: #ATTACHMENT }
      file_content          as FileContent,
      @Semantics.mimeType: true
      mime_type             as MimeType,
      file_name             as FileName,

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
