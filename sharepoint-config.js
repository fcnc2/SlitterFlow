// Fill with verified IDs and INTERNAL column names from the existing SharePoint lists.
// Never put PINs, passwords or client secrets here.
window.SLITTER_SHAREPOINT = {
  siteId: 'fitesacnc.sharepoint.com,9c41b493-3a6a-4c01-8d43-26b203783d58,1e97269c-5340-47e3-b05b-6170a179d56b',
  schemaVerified: true,
  lists: {
    ProductionRecords: {
      id: 'cb93c4d0-9fe2-4a54-b867-409a96bf1e0a',
      fields: { docNo:'Title', revision:'Revision', isLatest:'IsLatest', recordDate:'RecordDate', shift:'Shift', team:'Team', operatorId:'OperatorID', productCode:'ProductCode', jumboNo:'JumboNo', conditionRound:'ConditionRound', slitNo:'SlitNo', width:'WidthOnJumbo', speed:'Speed', ramp:'RampUpDown', tension:'WinderTension', pressure:'RiderRollPressure', torque:'Torque', rollLength:'RollLength', rollDiameter:'RollDiameter', remark:'Remark' }
    },
    MachineInspections: {
      id: 'ba3a8dbd-62e6-48d4-8a6e-a6473ce31145',
      fields: { inspectionKey:'Title', revision:'Revision', inspectedAt:'InspectionDateTime', docNo:'DocumentNo', shift:'Shift', team:'Team', operatorId:'OperatorID', area:'Area', overallResult:'OverallResult', remark:'Remark' }
    },
    InspectionResults: {
      id: '33f6a027-9be5-41d8-b231-05393d8850b2',
      fields: { resultKey:'Title', inspectionKey:'ColumnType_x002f__x0e15__x0e31__', docNo:'DocumentNo', revision:'Revision', area:'Area', category:'Category', itemCode:'ItemCode', checklistName:'ChecklistName', inputType:'InputType', resultText:'ResultText', resultNumber:'ResultNumber', resultStatus:'ResultStatus', operatorId:'OperatorID', checkedAt:'CheckedDateTime', remark:'Remark' }
    },
    ConditionValues: {
      id: '17479d1f-056a-4783-a311-ac62613d9ba6',
      fields: { title:'Title', recordKey:'RecordKey', docNo:'DocumentNo', revision:'Revision', area:'Area', productCode:'ProductCode', jumboNo:'JumboNo', conditionRound:'ConditionRound', parameterCode:'ParameterCode', parameterName:'ParameterName', dataType:'DataType', valueNumber:'ValueNumber', valueText:'ValueText', unit:'Unit', operatorId:'OperatorID', recordedAt:'RecordedDateTime', remark:'Remark' }
    },
    KnifeSelectionRecords: {
      id: 'd0c5e332-16b1-4fed-afd3-8c20d16cd38a',
      fields: { title:'Title', recordKey:'RecordKey', docNo:'DocumentNo', revision:'Revision', area:'Area', productCode:'ProductCode', jumboNo:'JumboNo', widthPattern:'WidthPattern', knifeNo:'KnifeNo', operatorId:'OperatorID', recordedAt:'RecordedDateTime', remark:'Remark' }
    },
    KnifeReplacements: {
      id: '925b3412-b79f-4815-9473-27790ac18ac7',
      fields: { title:'Title', recordKey:'RecordKey', docNo:'DocumentNo', revision:'Revision', area:'Area', productCode:'ProductCode', jumboNo:'JumboNo', knifeNo:'KnifeNo', lifeHours:'UsageLife', lifeUnit:'UsageLifeUnit', reason:'Reason', replacedAt:'ReplacementDateTime', operatorId:'OperatorID', remark:'Remark' }
    }
  }
};
