// Fill with verified IDs and INTERNAL column names from the existing SharePoint lists.
// Never put PINs, passwords or client secrets here.
window.SLITTER_SHAREPOINT = {
  siteId: '',
  schemaVerified: false,
  lists: {
    SlitterRecord: { id: '', fields: {} },
    PreStartCheck: { id: '', fields: {} },
    SlitCondition: { id: '', fields: {} },
    KnifeUsage: { id: '', fields: {} },
    KnifeReplacement: { id: '', fields: {} }
  }
};
