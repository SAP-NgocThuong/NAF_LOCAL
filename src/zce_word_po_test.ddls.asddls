@EndUserText.label: 'Custom entity Word PO test'
@ObjectModel.query.implementedBy: 'ABAP:ZCL_WORD_PO_QUERY_TEST'
@UI.headerInfo: { typeName:       'Đơn mua hàng',
                  typeNamePlural: 'Xuất Word đơn mua hàng (thử)',
                  title:       { type: #STANDARD, value: 'PurchaseOrder' },
                  description: { type: #STANDARD, value: 'SupplierName' } }
define custom entity ZCE_WORD_PO_TEST
{
      @UI.facet: [ { id: 'general', purpose: #STANDARD, type: #IDENTIFICATION_REFERENCE,
                     label: 'Thông tin', position: 10 } ]
      @EndUserText.label: 'Đơn mua hàng'
      @UI.lineItem:       [ { position: 10 } ]
      @UI.identification: [ { position: 10 } ]
      @UI.selectionField: [ { position: 10 } ]
  key PurchaseOrder     : abap.char(10);

      @EndUserText.label: 'Company Code'
      @UI.lineItem:       [ { position: 20 } ]
      @UI.identification: [ { position: 20 } ]
      @UI.selectionField: [ { position: 20 } ]
      CompanyCode       : abap.char(4);

      @EndUserText.label: 'Nhà cung cấp'
      @UI.lineItem:       [ { position: 30 } ]
      @UI.identification: [ { position: 30 } ]
      @UI.selectionField: [ { position: 30 } ]
      Supplier          : abap.char(10);

      @EndUserText.label: 'Tên nhà cung cấp'
      @UI.lineItem:       [ { position: 40 } ]
      @UI.identification: [ { position: 40 } ]
      @Consumption.filter.hidden: true
      SupplierName      : abap.char(80);

      @EndUserText.label: 'Ngày đơn'
      @UI.lineItem:       [ { position: 50 } ]
      @UI.identification: [ { position: 50 } ]
      @UI.selectionField: [ { position: 40 } ]
      PurchaseOrderDate : abap.dats;

      @EndUserText.label: 'Tiền tệ'
      @UI.lineItem:       [ { position: 60 } ]
      @Consumption.filter.hidden: true
      DocumentCurrency  : abap.cuky;

      @EndUserText.label: 'File Word'
      @UI.identification: [ { position: 60 } ]
      @Consumption.filter.hidden: true
      @Semantics.largeObject: { mimeType: 'FileMime',
                                fileName: 'FileName',
                                contentDispositionPreference: #ATTACHMENT }
      FileContent       : abap.rawstring(0);

      @UI.hidden: true
      @Consumption.filter.hidden: true
      @Semantics.mimeType: true
      FileMime          : abap.char(128);

      @UI.hidden: true
      @Consumption.filter.hidden: true
      FileName          : abap.char(128);

      @EndUserText.label: 'Thông báo'
      @UI.identification: [ { position: 70 } ]
      @Consumption.filter.hidden: true
      Message           : abap.char(255);
}
