Sub Auto_Open()
UserForm8.Show
End Sub
Sub puntitos()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("Resumen Puntos").Select
Range("DH5").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(Liquidacion!R2C9,Meses!C1:C3,3,FALSE)"
Range("CP1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C1)-1"
vari_9 = Range("CP1").Value
Range("A6").Select

For i = 1 To vari_9
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-1],Liquidacion!C3:C30,28,FALSE),IFERROR(VLOOKUP('Resumen Puntos'!RC[-1],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C30,28,FALSE),""CEDULA NO ENCONTRADA""))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-2],Liquidacion!C3:C14,12,FALSE),IFERROR(VLOOKUP(RC[-2],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C14,12,FALSE),""CÉDULA NO ENCONTRADA""))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-3],Liquidacion!C3:C26,24,FALSE),IFERROR(VLOOKUP('Resumen Puntos'!RC[-3],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C26,24,FALSE),""CÉDULA NO ENCONTRADA""))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-4],Liquidacion!C3:C20,16,FALSE),IFERROR(VLOOKUP('Resumen Puntos'!RC[-4],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C20,16,FALSE),""CÉDULA NO ENCONTRADA""))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5],MOVIL!C1:C26,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-3]>=R3C4,RC[-1]*VLOOKUP(RC[-3],'Tabla Var'!R18C15:R19C17,3,TRUE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-7],MOVIL!C1:C26,3,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-5]>=R3C4,RC[-1]*VLOOKUP('Resumen Puntos'!RC[-5],'Tabla Var'!R18C15:R19C18,4,TRUE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-9],MOVIL!C1:C26,4,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-7]>=R3C4,RC[-1]*VLOOKUP('Resumen Puntos'!RC[-7],'Tabla Var'!R18C15:R19C19,5,TRUE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-11],MOVIL!C1:C6,6,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-10]>=R3C3,RC[-9]>=R3C4),RC[-1]*VLOOKUP(RC[-9],'Tabla Var'!R18C6:R23C8,3,TRUE),IF(AND(RC[-9]>=R3C4,RC[-10]<R3C3),RC[-1]*VLOOKUP(RC[-9],'Tabla Var'!R18C15:R19C17,3,TRUE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-13],MOVIL!C1:C7,7,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-12]>=R3C3,RC[-11]>=R3C4),RC[-1]*VLOOKUP(RC[-11],'Tabla Var'!R18C6:R23C9,4,TRUE),IF(AND('Resumen Puntos'!RC[-11]>=R3C4,'Resumen Puntos'!RC[-12]<R3C3),'Resumen Puntos'!RC[-1]*VLOOKUP('Resumen Puntos'!RC[-11],'Tabla Var'!R18C15:R19C18,4,TRUE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-15],MOVIL!C1:C8,8,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-14]>=R3C3,RC[-13]>=R3C4),RC[-1]*VLOOKUP(RC[-13],'Tabla Var'!R18C6:R23C10,5,TRUE),IF(AND('Resumen Puntos'!RC[-13]>=R3C4,'Resumen Puntos'!RC[-14]<R3C3),'Resumen Puntos'!RC[-1]*VLOOKUP('Resumen Puntos'!RC[-13],'Tabla Var'!R18C15:R19C19,5,TRUE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-1],RC[-3],RC[-5],RC[-7],RC[-9],RC[-11])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(VLOOKUP(RC[-18],MOVIL!C1:C14,14,FALSE),0)>=50,0,IFERROR(VLOOKUP(RC[-18],MOVIL!C1:C26,11,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(VLOOKUP(RC[-19],MOVIL!C1:C14,14,FALSE),0)>=50,0,IFERROR(VLOOKUP(RC[-19],MOVIL!C1:C26,12,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(VLOOKUP(RC[-20],MOVIL!C1:C14,14,FALSE),0)>=50,0,IFERROR(VLOOKUP(RC[-20],MOVIL!C1:C26,13,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-19]>=R3C3,RC[-18]>=R3C4),RC[1]*VLOOKUP(RC[-18],'Tabla Var'!R18C6:R23C8,3,TRUE),IF(AND(RC[-18]>=R3C4,RC[-19]<R3C3),RC[1]*VLOOKUP(RC[-18],'Tabla Var'!R18C15:R19C17,3,TRUE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-21]>=R3C3,RC[-20]>=R3C4),RC[1]*VLOOKUP(RC[-20],'Tabla Var'!R18C6:R23C9,4,TRUE),IF(AND('Resumen Puntos'!RC[-20]>=R3C4,'Resumen Puntos'!RC[-21]<R3C3),'Resumen Puntos'!RC[1]*VLOOKUP('Resumen Puntos'!RC[-20],'Tabla Var'!R18C15:R19C18,4,TRUE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-5]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-23]>=R3C3,RC[-22]>=R3C4),RC[1]*VLOOKUP(RC[-22],'Tabla Var'!R18C6:R23C10,5,TRUE),IF(AND('Resumen Puntos'!RC[-22]>=R3C4,'Resumen Puntos'!RC[-23]<R3C3),'Resumen Puntos'!RC[1]*VLOOKUP('Resumen Puntos'!RC[-22],'Tabla Var'!R18C15:R19C19,5,TRUE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-6]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-2],RC[-4],RC[-6],RC[3])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-25]>=R3C4,IFERROR(VLOOKUP(RC[-28],MOVIL!C1:C35,35,FALSE),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-26]>=R3C4,IFERROR(VLOOKUP(RC[-29],MOVIL!C1:C34,34,FALSE),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
      ActiveCell.FormulaR1C1 = _
        "=IF(RC[-27]>=R3C4,IF(IFERROR(VLOOKUP(RC[-30],MOVIL!C1:C14,14,FALSE),0)>=50,IFERROR(VLOOKUP(RC[-30],MOVIL!C1:C36,36,FALSE),0),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-28]<R3C4,0,IF(AND(RC[-28]>=R3C4,RC[-27]>=R3C5,RC[-29]>=R3C3,RC[-30]>=R3C2),RC[1]*1.5%,RC[1]*0.7%))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=(IFERROR(VLOOKUP(RC[-32],MOVIL!C1:C30,27,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-30]<R3C4,0,IF(AND(RC[-30]>=R3C4,RC[-29]>=R3C5,RC[-31]>=R3C3,RC[-32]>=R3C2),RC[1]*1.5%,RC[1]*0.7%))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=(IFERROR(VLOOKUP(RC[-34],MOVIL!C1:C29,29,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-32]<R3C4,0,IF(AND(RC[-32]>=R3C4,RC[-31]>=R3C5,RC[-33]>=R3C3,RC[-34]>=R3C2),RC[1]*1.5%,RC[1]*0.7%))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=(IFERROR(VLOOKUP(RC[-36],MOVIL!C1:C28,28,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-34]<R3C4,0,IF(AND(RC[-34]>=R3C4,RC[-33]>=R3C5,RC[-35]>=R3C3,RC[-36]>=R3C2),RC[1]*1.5%,RC[1]*0.7%))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=(IFERROR(VLOOKUP(RC[-38],MOVIL!C1:C30,30,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-36]>=R3C4,IFERROR(VLOOKUP(RC[-39],'Fuente Hogares'!C12:C13,2,FALSE)*5000,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-40],'Fuente Hogares'!C12:C13,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-38]>=R3C4,IFERROR(VLOOKUP(RC[-41],'Fuente Hogares'!C4:C5,2,FALSE)*5000,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-42],'Fuente Hogares'!C4:C5,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-40]>=R3C4,RC[1]*5000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(SUM(IFERROR(VLOOKUP(RC[-44],MOVIL!C1:C22,17,FALSE),0),IFERROR(VLOOKUP(RC[-44],MOVIL!C1:C22,18,FALSE),0)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-42]>=R3C4,IFERROR((IFERROR(VLOOKUP(RC[-45],MOVIL!C1:C26,19,FALSE),0))*5500,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR((IFERROR(VLOOKUP(RC[-46],MOVIL!C1:C26,19,FALSE),0)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-44]>=R3C4,IFERROR((IFERROR(VLOOKUP(RC[-47],MOVIL!C1:C26,20,FALSE),0))*11000,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR((IFERROR(VLOOKUP(RC[-48],MOVIL!C1:C26,20,FALSE),0)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-46]>=R3C4,IFERROR((IFERROR(VLOOKUP(RC[-49],MOVIL!C1:C26,21,FALSE),0))*22000,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR((IFERROR(VLOOKUP(RC[-50],MOVIL!C1:C26,21,FALSE),0)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-48]>=R3C4,IFERROR((VLOOKUP(IFERROR(VLOOKUP(RC[-51],MOVIL!C1:C26,22,FALSE),0),'Tabla Var'!R3C9:R7C11,3,TRUE))*(IFERROR(VLOOKUP(RC[-51],MOVIL!C1:C26,22,FALSE),0)),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-52],MOVIL!C1:C26,22,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-50]>=R3C4,(((IFERROR(VLOOKUP(RC[-53],MOVIL!C1:C26,23,FALSE),0)))*1700),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=(IFERROR(VLOOKUP(RC[-54],MOVIL!C1:C26,23,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-52]>=R3C4,(((IFERROR(VLOOKUP(RC[-55],MOVIL!C1:C26,24,FALSE),0)))*3900),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-56],MOVIL!C1:C26,24,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-55]>=R3C3,RC[-54]>=R3C4,RC[-54]>=100%,RC[37]>=R3C95),RC[1]*VLOOKUP(RC[-54],'Tabla Var'!R18C6:R23C11,6,TRUE),IF(AND(RC[-55]>=R3C3,RC[-54]>=R3C4,RC[-54]>=100%,RC[37]<R3C95),RC[1]*12000,IF(AND(RC[-55]>=R3C3,RC[-54]>=R3C4,RC[-54]<100%),RC[1]*VLOOKUP(RC[-54],'Tabla Var'!R18C6:R23C11,6,TRUE),IF(AND('Resumen Puntos'!RC[-54]>=R3C4,'Resumen Puntos'!RC[-55]<R3C3),'Resumen Puntos'!RC[1]*VLOOKUP('Resumen Puntos'!RC[-54],'Tabla Var'!R18C15:R19C20,6,TRUE),0))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-58],'Fuente Hogares'!C78:C79,2,FALSE),0)"
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-58],'Fuente Hogares'!C78:C79,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-56]>=R3C4,RC[1]*VLOOKUP('Resumen Puntos'!RC[-56],'Tabla Var'!R18C15:R19C20,6,TRUE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-60],'Fuente Hogares'!C1:C2,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-59]>=R3C3,RC[-58]>=R3C4,RC[-58]>=100%,RC[33]>=R3C95),RC[1]*VLOOKUP(RC[-58],'Tabla Var'!R18C6:R23C11,6,TRUE),IF(AND(RC[-59]>=R3C3,RC[-58]>=R3C4,RC[-58]>=100%,RC[33]<R3C95),RC[1]*12000,IF(AND(RC[-59]>=R3C3,RC[-58]>=R3C4,RC[-58]<100%),RC[1]*VLOOKUP(RC[-58],'Tabla Var'!R18C6:R23C11,6,TRUE),IF(AND('Resumen Puntos'!RC[-58]>=R3C4,'Resumen Puntos'!RC[-59]<R3C3),'Resumen Puntos'!RC[1]*VLOOKUP('Resumen Puntos'!RC[-58],'Tabla Var'!R18C15:R19C20,6,TRUE),0))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-62],'Fuente Hogares'!C82:C83,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-60]<R3C4,0,IF(AND(RC[-60]>=R3C4,RC[-59]>=R3C5,RC[-61]>=R3C3,RC[-62]>=R3C2),RC[1]*1.5%,RC[1]*0.7%))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(IFERROR(VLOOKUP(RC[-64],'Fuente Hogares'!C57:C58,2,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-62]<R3C4,0,IF(AND(RC[-62]>=R3C4,RC[-61]>=R3C5,RC[-63]>=R3C3,RC[-64]>=R3C2),RC[1]*1.5%,RC[1]*0.7%))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(IFERROR(VLOOKUP(RC[-66],'Fuente Hogares'!C60:C61,2,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-64]>=R3C4,IFERROR(VLOOKUP(RC[-67],'Fuente Hogares'!C16:C17,2,0)*5500,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-68],'Fuente Hogares'!C16:C17,2,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-66]>=R3C4,((IFERROR(VLOOKUP(RC[-69],'Fuente Hogares'!C19:C20,2,FALSE),0)*11000)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-70],'Fuente Hogares'!C19:C20,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-68]>=R3C4,((IFERROR(VLOOKUP(RC[-71],'Fuente Hogares'!C22:C23,2,FALSE),0)*22000)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-72],'Fuente Hogares'!C22:C23,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-70]>=R3C4,((IFERROR(VLOOKUP(RC[-73],'Fuente Hogares'!C36:C37,2,FALSE),0))*2000),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-74],'Fuente Hogares'!C36:C37,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-72]>=R3C4,((IFERROR(VLOOKUP(RC[-75],'Fuente Hogares'!C46:C47,2,FALSE),0)*5000)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-76],'Fuente Hogares'!C46:C47,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-74]>=R3C4,IFERROR(VLOOKUP(RC[-77],MOVIL!C1:C26,26,FALSE),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-75]>=R3C4,IF(IFERROR(VLOOKUP(RC[-78],'Fuente Hogares'!C74:C75,2,FALSE),0)<=10,(IFERROR(VLOOKUP(RC[-78],'Fuente Hogares'!C74:C75,2,FALSE),0)*2000),(IFERROR(VLOOKUP(RC[-78],'Fuente Hogares'!C74:C75,2,FALSE),0)*3000)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-79],'Fuente Hogares'!C74:C75,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-77]>=R3C4,((IFERROR(VLOOKUP(RC[-80],'Fuente Hogares'!C55:C56,2,FALSE),0)*5000)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-81],'Fuente Hogares'!C55:C56,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-79]>=R3C4,((IFERROR(VLOOKUP(RC[-82],'Fuente Hogares'!C52:C53,2,FALSE),0)*2200)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-83],'Fuente Hogares'!C52:C53,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-81]>=R3C4,RC[1]*30%,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-85],MOVIL!C1:C33,33,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-83]>=R3C4,((IFERROR(VLOOKUP(RC[-86],'Fuente Hogares'!C64:C65,2,FALSE),0)*2200)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-87],'Fuente Hogares'!C64:C65,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-85]>=R3C4,RC[1]*10%,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-89],MOVIL!C1:C38,38,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-87]>=R3C4,(IFERROR(VLOOKUP(RC[-90],'Fuente Hogares'!C71:C72,2,FALSE),0)*2200),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-91],'Fuente Hogares'!C71:C72,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-89]>=R3C4,(IFERROR(VLOOKUP(RC[-92],'Fuente Hogares'!C67:C68,2,FALSE),0)*5000),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-93],'Fuente Hogares'!C67:C68,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-94],Liquidacion!C3:C91,89,FALSE),IFERROR(VLOOKUP(RC[-94],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C78,76,FALSE),""CÉDULA NO ENCONTRADA""))"
    ActiveCell.Offset(0, 5).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(RC[-11],RC[-15],RC[-34],RC[-36],RC[-38],RC[-40],RC[-42],RC[-62],RC[-64],RC[-66],RC[-68],RC[-71]:RC[-70],RC[-72],RC[-82])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[1]=0,RC[7]=0),IF(AND(RC[6]>=100%,RC[10]>=100%),SUM(RC[-8],RC[-10],RC[-14],RC[-18],RC[-20],RC[-22],RC[-23],RC[-25],RC[-27],RC[-29],RC[-31],RC[-33],RC[-45],RC[-47],RC[-49],RC[-51],RC[-53],RC[-55],RC[-57],RC[-59],RC[-61])*120%,IF(AND(RC[6]>=100%,RC[10]<100%),SUM(RC[-8],RC[-10],RC[-14],RC[-18],RC[-20],RC[-22],RC[-23],RC[-25],RC[-27],RC[-29],RC[-31],RC[-33],RC[-45],RC[-47],RC[-49],RC[-51],RC[-53],RC[-55],RC[-57],RC[-59],RC[-61])*90%,IF(AND(RC[6]<100%,RC[10]>=100%),SUM(RC[-8],RC[-10],RC[-14],RC[-18],RC[-20],RC[-22],RC[-23],RC[-25],RC[-27],RC[-29],RC[-31],RC[-33],RC[-45],RC[-47],RC[-49],RC[-51],RC[-53],RC[-55],RC[-57],RC[-59],RC[-61])*80%,IF(AND(RC[6]<100%,RC[10]<100%),0,0)))),IF(AND(RC[3]>=100%,RC[6]>=100%,RC[10]>=100%)," & _
        "SUM(RC[-8],RC[-10],RC[-14],RC[-18],RC[-20],RC[-22],RC[-23],RC[-25],RC[-27],RC[-29],RC[-31],RC[-33],RC[-45],RC[-47],RC[-49],RC[-51],RC[-53],RC[-55],RC[-57],RC[-59],RC[-61])*150%,IF(AND(RC[3]>=100%,RC[6]>=100%,RC[10]<100%),SUM(RC[-8],RC[-10],RC[-14],RC[-18],RC[-20],RC[-22],RC[-23],RC[-25],RC[-27],RC[-29],RC[-31],RC[-33],RC[-45],RC[-47],RC[-49],RC[-51],RC[-53],RC[-55],RC[-57],RC[-59],RC[-61])*120%,IF(AND(RC[3]>=100%,RC[6]<100%,RC[10]>=100%),SUM(RC[-8],RC[-10],RC[-14],RC[-18],RC[-20],RC[-22],RC[-23],RC[-25],RC[-27],RC[-29],RC[-31],RC[-33],RC[-45],RC[-47],RC[-49],RC[-51],RC[-53],RC[-55],RC[-57],RC[-59],RC[-61])*110%,IF(AND(RC[3]<100%,RC[6]>=100%,RC[10]>=100%),SUM(RC[-8],RC[-10],RC[-14],RC[-18],RC[-20],RC[-22],RC[-23],RC[-25],RC[-27],RC[-29],RC[-31],RC[-33]," & _
        "RC[-45],RC[-47],RC[-49],RC[-51],RC[-53],RC[-55],RC[-57],RC[-59],RC[-61]),IF(AND(RC[3]>=100%,RC[6]<100%,RC[10]<100%),SUM(RC[-8],RC[-10],RC[-14],RC[-18],RC[-20],RC[-22],RC[-23],RC[-25],RC[-27],RC[-29],RC[-31],RC[-33],RC[-45],RC[-47],RC[-49],RC[-51],RC[-53],RC[-55],RC[-57],RC[-59],RC[-61])*90%,IF(AND(RC[3]<100%,RC[6]>=100%,RC[10]<100%),SUM(RC[-8],RC[-10],RC[-14],RC[-18],RC[-20],RC[-22],RC[-23],RC[-25],RC[-27],RC[-29],RC[-31],RC[-33],RC[-45],RC[-47],RC[-49],RC[-51],RC[-53],RC[-55],RC[-57],RC[-59],RC[-61])*80%,IF(AND(RC[3]<100%,RC[6]<100%,OR(RC[10]>=100%,RC[10]<100%)),0,0))))))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-101],Metas!C41:C43,3,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-1]=0,IFERROR(VLOOKUP(RC[-102],Liquidacion!C3:C10,8,FALSE),VLOOKUP(RC[-102],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C10,8,FALSE))<=0),0,IF(AND(ROUND(RC[-1]*(IFERROR(VLOOKUP(RC[-102],Liquidacion!C3:C10,8,FALSE),VLOOKUP(RC[-102],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C10,8,FALSE))/R5C112),0)>=0,ROUND(RC[-1]*(IFERROR(VLOOKUP(RC[-102],Liquidacion!C3:C10,8,FALSE),VLOOKUP(RC[-102],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C10,8,FALSE))/R5C112),0)<=0.5),1,ROUND(RC[-1]*(IFERROR(VLOOKUP(RC[-102],Liquidacion!C3:C10,8,FALSE),VLOOKUP(RC[-102],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C10,8,FALSE))/R5C112),0)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-27]/RC[-1],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-104],Metas!C41:C42,2,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-105],'NPS-UMBRAL'!C55:C56,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-107],Metas!C41:C47,7,FALSE),""Sin metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(VLOOKUP(RC[-108],Liquidacion!C3:C10,8,FALSE),VLOOKUP(RC[-108],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C10,8,FALSE))<=0,0,IF(AND(ROUND(RC[-1]*(IFERROR(VLOOKUP(RC[-108],Liquidacion!C3:C10,8,FALSE),VLOOKUP(RC[-108],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C10,8,FALSE))/R5C112),0)>=0,ROUND(RC[-1]*(IFERROR(VLOOKUP(RC[-108],Liquidacion!C3:C10,8,FALSE),VLOOKUP(RC[-108],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C10,8,FALSE))/R5C112),0)<=0.5),1,ROUND(RC[-1]*(IFERROR(VLOOKUP(RC[-108],Liquidacion!C3:C10,8,FALSE),VLOOKUP(RC[-108],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C10,8,FALSE))/R5C112),0)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(RC[-69],RC[-67],RC[-35],RC[-28],RC[-26],RC[-18],RC[-16])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlToLeft).Select

Next i

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation


End Sub

Sub liquida_trunks()
'

'

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Dim variable_1 As Object
Sheets("Liquidacion").Select
Range("I2") = UserForm7.ComboBox1
Range("a1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C[2])-2"
vari_9 = Range("a1").Value
Range("c4").Select

For i = 1 To vari_9

ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],HC!C3:C5,3,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],HC!C3:C7,5,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-3],Metas!C1:C5,3,0),""Sin Oficina"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-4],Metas!C1:C5,5,0),""Sin Oficina"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(VLOOKUP(RC[-5],HC!C3:C9,7,0),0)>0,VLOOKUP(RC[-5],'DESARROLLO+PROYECTOS'!C3:C9,7,FALSE),IFERROR(VLOOKUP(RC[-5],HC!C3:C9,7,0),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C9,Meses!C[-8]:C[-7],2,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
   ActiveCell.FormulaR1C1 = _
        "=IF((IF((IF(AND(RC[69]=""retiro"",RC[66]=RC[64],RC[67]=RC[-1]),IF((IF(RC[65]>(CONCATENATE(RC[64],RC[-1],RC[63])-RC[61]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[63]-RC[68])-RC[61])+1,(RC[63]-RC[68]+1)+(RC[63]-(RC[61]+RC[62]))),IF(AND(RC[66]=RC[64],RC[67]=RC[-1]),(RC[63]-RC[68]+1)-(RC[61]+RC[62]),RC[63]-(RC[61]+RC[62]))))<0,0,IF(AND(RC[69]=""retiro"",RC[66]=RC[64],RC[67]=RC[-1]),IF((IF(RC[65]>(CONCATENATE(RC[64],RC[-1],RC[63])-RC[61]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[63]-RC[68])-RC[61])+1,(RC[63]-RC[68]+1)+(RC[63]-(RC[61]+RC[62]))),IF(AND(RC[66]=RC[64],RC[67]=RC[-1]),(RC[63]-RC[68]+1)-(RC[61]+RC[62]),RC[63]-(RC[61]+RC[62])))-IFERROR(VLOOKUP(RC[-7],'AISLAMIENTO COVID'!C1:C7,7,FALSE),0)))<0,0," & _
        "IF((IF(AND(RC[69]=""retiro"",RC[66]=RC[64],RC[67]=RC[-1]),IF((IF(RC[65]>(CONCATENATE(RC[64],RC[-1],RC[63])-RC[61]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[63]-RC[68])-RC[61])+1,(RC[63]-RC[68]+1)+(RC[63]-(RC[61]+RC[62]))),IF(AND(RC[66]=RC[64],RC[67]=RC[-1]),(RC[63]-RC[68]+1)-(RC[61]+RC[62]),RC[63]-(RC[61]+RC[62]))))<0,0,IF(AND(RC[69]=""retiro"",RC[66]=RC[64],RC[67]=RC[-1]),IF((IF(RC[65]>(CONCATENATE(RC[64],RC[-1],RC[63])-RC[61]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[63]-RC[68])-RC[61])+1,(RC[63]-RC[68]+1)+(RC[63]-(RC[61]+RC[62]))),IF(AND(RC[66]=RC[64],RC[67]=RC[-1]),(RC[63]-RC[68]+1)-(RC[61]+RC[62]),RC[63]-(RC[61]+RC[62])))-IFERROR(VLOOKUP(RC[-7],'AISLAMIENTO COVID'!C1:C7,7,FALSE),0)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-8],Metas!C1:C26,13,0),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-2]<=0,0,IF(OR(RC[-1]=0,RC[-1]=""Sin Metas""),0,IF((ROUND(IFERROR(VLOOKUP(RC[-9],Metas!C1:C25,13,0)*(RC[-2]/RC[61]),""Sin Metas""),0))<=0,1,(ROUND(IFERROR(VLOOKUP(RC[-9],Metas!C1:C25,13,0)*(RC[-2]/RC[61]),""Sin Metas""),0)))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(IFERROR(VLOOKUP(RC[-10],'Fuente Hogares'!C1:C2,2,0),0),IFERROR(VLOOKUP(RC[-10],'Fuente Hogares'!C78:C79,2,0),0),IFERROR(VLOOKUP(RC[-10],'Fuente Hogares'!C82:C83,2,0),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-4]=0,0,IFERROR(IF(AND(RC[-2]=0,RC[-1]>=1),100%,IF(AND(RC[-2]=0,RC[-1]=0),100%,RC[-1]/RC[-2])),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-12],Metas!C1:C26,12,0),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(RC[-6]<=0,0,IF(RC[-1]=0,0,IF((ROUND(IFERROR(VLOOKUP(RC[-13],Metas!C1:C25,12,0)*(RC[-6]/RC[57]),""Sin Metas""),0))<=0,1,(ROUND(IFERROR(VLOOKUP(RC[-13],Metas!C1:C25,12,0)*(RC[-6]/RC[57]),""Sin Metas""),0))))),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(IFERROR(VLOOKUP(RC[-14],MOVIL!C1:C10,10,FALSE),0),IFERROR(VLOOKUP(RC[-14],MOVIL!C1:C14,14,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-8]=0,0,IFERROR(IF(RC[-2]=0,100%,RC[-1]/RC[-2]),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(RC[-2],IF(AND(RC[-1]>=90%,RC[5]>=80%),SUM(VLOOKUP(RC[-16],Resumen_plan_power!C1:C5,5,FALSE),VLOOKUP(RC[-16],Resumen_plan_power!C1:C7,7,FALSE),VLOOKUP(RC[-16],Resumen_plan_power!C1:C17,17,FALSE)),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-10]=0,0,IFERROR(IF(RC[-4]=0,100%,RC[-1]/RC[-4]),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-6]=""Sin Metas"",""Sin Metas"",SUM(RC[-6],RC[-10]))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=+RC[-10]+RC[-6]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-6]+RC[-10]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-14]=0,0,IFERROR(IF(AND(RC[-2]=0,RC[-1]=0),0%,RC[-1]/RC[-2]),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-6]+RC[-12]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-16]=0,0,IFERROR(IF(AND(RC[-4]=0,RC[-1]=0),0%,RC[-1]/RC[-4]),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-24],Metas!C[-26]:C[-2],14,FALSE),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-18]<=0,0,IFERROR(RC[-1]*(RC[-18]/RC[45]),""Sin Metas""))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(IFERROR(VLOOKUP(RC[-26],MOVIL!C1:C15,15,FALSE),0),IFERROR(VLOOKUP(RC[-26],MOVIL!C1:C16,16,FALSE),0),IFERROR(VLOOKUP(RC[-26],'Fuente Hogares'!C7:C8,2,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-20]=0,0,IF(RC[-2]=0,0,IF(AND(RC[-2]=0,RC[-1]>=1),100%,IF(AND(RC[-2]=0,RC[-1]=0),0%,IFERROR(IF(AND(RC[-2]=0,RC[-1]=0),0%,RC[-1]/RC[-2]),0)))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-28],'Resumen Puntos'!C1:C100,100,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-29],'Resumen Puntos'!C1:C101,101,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-25]=0,RC[-7]<80%),0,IF(AND(RC[-25]=0,RC[-7]>160%),160%*55%,IF(RC[-25]=0,RC[-7]*55%,IFERROR(RC[-2]/RC[-25],0))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-26]=0,0,RC[-2]+RC[-3])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-32],Metas!C1:C8,8,0),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-26]<=0,0,IFERROR(VLOOKUP(RC[-33],Metas!C1:C8,8,0)*(RC[-26]/RC[37]),""Sin Metas""))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(SUM(IFERROR(VLOOKUP(RC[-34],'Fuente Hogares'!C26:C27,2,FALSE),0),IFERROR(VLOOKUP(Liquidacion!RC[-34],'Fuente Hogares'!C29:C30,2,FALSE),0),0)) - (SUM(IFERROR(VLOOKUP(RC[-34],'Fuente Hogares'!C40:C41,2,FALSE),0),IFERROR(VLOOKUP(Liquidacion!RC[-34],'Fuente Hogares'!C43:C44,2,FALSE),0)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-3]=""Sin Metas"",0%,IF(RC[-28]=0,0,IF(AND(RC[-2]<0,RC[-1]<0),RC[-2]/RC[-1],IF(AND(RC[-2]<0,RC[-1]>=0),140%,IF(AND(RC[-2]>0,RC[-1]>0),RC[-1]/RC[-2],IF(AND(RC[-2]>0,RC[-1]<=0),0%,0%))))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(LOOKUP(RC[-1],'Tabla Var'!R3C13:R13C15)=""lineal"",RC[-1],LOOKUP(RC[-1],'Tabla Var'!R2C13:R13C15))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[14]=""Sin Metas"",RC[14]=""N/A"",RC[14]=0),RC[-1]*25%,RC[-1]*R2C40)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(RC[-1]*((SUM(RC[-31],IF(IFERROR(VLOOKUP(RC[-38],'NPS-UMBRAL'!C1:C10,10,FALSE),0)>8,IFERROR(VLOOKUP(RC[-38],'NPS-UMBRAL'!C1:C10,10,FALSE),0),0))/RC[32]))*RC[-33])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-39],Metas!C41:C44,4,FALSE),""Sin metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-40],'NPS-UMBRAL'!C1:C5,5,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-2]/RC[-1],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-42],Metas!C41:C45,5,FALSE),""Sin metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-43],'NPS-UMBRAL'!C1:C6,6,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-45],Metas!C1:C26,15,0),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(VLOOKUP(RC[-46],'NPS-UMBRAL'!C1:C17,15,0),0)=""NO REGISTRA"",""NO REGISTRA ENCUESTAS"",IFERROR(VLOOKUP(RC[-46],'NPS-UMBRAL'!C1:C17,15,0),""NO REGISTRA ENCUESTAS""))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-40]=0,0%,IFERROR(IF(RC[-1]<=0%,0%,RC[-1]/RC[-2]),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-7]>=100%,RC[-4]>=100%),IF(RC[36]>=100%,IF(RC[-2]=100%,160%,IF(VLOOKUP(RC[-1],'Tabla Var'!R2C1:R13C3,3,TRUE)=""lineal"",Liquidacion!RC[-1],VLOOKUP(RC[-1],'Tabla Var'!R2C1:R13C3,3,TRUE))),IF(VLOOKUP(RC[-1],'Tabla Var'!R2C1:R13C4,4,TRUE)=""lineal"",Liquidacion!RC[-1],VLOOKUP(RC[-1],'Tabla Var'!R2C1:R13C4,4,TRUE))),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]*R2C52,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(RC[-1]*((SUM(RC[-43],IF(IFERROR(VLOOKUP(RC[-50],'NPS-UMBRAL'!C1:C10,10,FALSE),0)>8,IFERROR(VLOOKUP(RC[-50],'NPS-UMBRAL'!C1:C10,10,FALSE),0),0))/RC[20]))*RC[-45])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-51],Metas!C1:C26,19,0),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-45]<=0,0,IF(OR(RC[-1]=""Sin Metas"",RC[-1]=0,RC[-1]=""N/A""),0,IF(AND((ROUND(IFERROR(VLOOKUP(RC[-52],Metas!C[-54]:C[-29],19,0)*(RC[-45]/RC[18]),""Sin Metas""),0))>=0,(ROUND(IFERROR(VLOOKUP(RC[-52],Metas!C[-54]:C[-29],19,0)*(RC[-45]/RC[18]),""Sin Metas""),0))<=0.5),1,ROUND(IFERROR(VLOOKUP(RC[-52],Metas!C[-54]:C[-29],19,0)*(RC[-45]/RC[18]),""Sin Metas""),0))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-53],MOVIL!C1:C25,25,FALSE),0) + IFERROR(VLOOKUP(RC[-53],'Fuente Hogares'!C33:C34,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-47]=0,0,IFERROR(RC[-1]/RC[-2],0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(LOOKUP(RC[-1],'Tabla Var'!R3C13:R13C15)=""lineal"",RC[-1],LOOKUP(RC[-1],'Tabla Var'!R2C13:R13C15))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]*R2C59,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(RC[-1]*((SUM(RC[-50],IF(IFERROR(VLOOKUP(RC[-57],'NPS-UMBRAL'!C1:C10,10,FALSE),0)>8,IFERROR(VLOOKUP(RC[-57],'NPS-UMBRAL'!C1:C10,10,FALSE),0),0))/RC[13]))*RC[-52])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-21]+RC[-9]+RC[-2]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((AND(RC[-25]=0,RC[-16]=0)),0,((RC[-1]*RC[-54])*(SUM(RC[-52],IF(IFERROR(VLOOKUP(RC[-59],'NPS-UMBRAL'!C1:C10,10,FALSE),0)>8,IFERROR(VLOOKUP(RC[-59],'NPS-UMBRAL'!C1:C10,10,FALSE),0),0))/RC[11])))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-30]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-54]=0,RC[-56]=0),0,VLOOKUP(RC[-61],Resumen_plan_power!C1:C15,15,FALSE))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM((IF(AND(RC[-28]=0,RC[-19]=0),0,SUM(RC[-3],RC[-31],RC[-1]))),RC[18])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[14]>0,RC[-1]>0,RC[-1]<RC[14]),0,IF(AND(RC[14]>0,RC[-1]>0,RC[-1]>RC[14]),RC[-1]-RC[14],RC[-1]))"
    ActiveCell.Offset(0, 5).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-68],'Ausentismos-Vaca-Umb'!C[-70]:C[-64],7,0),0)+IFERROR(IF(VLOOKUP(RC[-68],'NPS-UMBRAL'!C1:C11,10,0)>8,VLOOKUP(RC[-68],'NPS-UMBRAL'!C1:C11,10,0),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-69],'Ausentismos-Vaca-Umb'!C[-61]:C[-57],5,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C9,Meses!C[-72]:C[-70],3,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C9,Meses!C[-73]:C[-70],4,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-72],HC!C3:C6,4,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-1],1,4)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-2],5,2)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-3],7,2)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-76],'Ausentismos-Vaca-Umb'!C[-78]:C[-74],5,0),""S/N"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-77],Garantizado!C3:C5,3,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-78],HC!C3:C20,18,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-79],HC!C3:C12,10,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-73]=0,0,IFERROR(VLOOKUP(RC[-80],'DESARROLLO+PROYECTOS'!C3:C34,32,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-27],RC[-34],RC[-46],RC[-54],RC[-60])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-82],Metas!C41:C46,6,0),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "='Tabla Var'!R16C2"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-2]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-85],Metas!C41:C48,8,FALSE),""Sin metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-79]<=0,0,IF(OR(RC[-1]=0,RC[-1]=""Sin Metas""),0,IF((ROUND(IFERROR(VLOOKUP(RC[-86],Metas!C41:C48,8,0)*(RC[-79]/RC[-16]),""Sin Metas""),0))<=0,1,(ROUND(IFERROR(VLOOKUP(RC[-86],Metas!C41:C48,8,0)*(RC[-79]/RC[-16]),""Sin Metas""),0)))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-87],'Fuente Hogares'!C87:C88,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]=0,100%,IFERROR(RC[-1]/RC[-2],0))"
    ActiveCell.Offset(1, 0).Select
    Selection.End(xlToLeft).Select
    
    Next i
     
    
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
End Sub




Sub horas()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("Liquidacion_PART_TIME").Select
Range("I2") = UserForm7.ComboBox1
Range("a1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C[2])-2"
vari_9 = Range("a1").Value
Range("c4").Select

For i = 1 To vari_9

    ActiveCell.Offset(0, 1).Range("A1").Select
    On Error GoTo marcador
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],HC!C3:C5,3,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],HC!C3:C7,5,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-3],Metas!C1:C5,3,0),""Sin Oficina"")"
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5],HC!C3:C9,7,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C9,Meses!C[-8]:C[-7],2,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-7],'Ausentismos-Vaca-Umb'!C31:C37,5,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-8],Metas!C1:C26,13,0),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=ROUND(IFERROR((VLOOKUP(RC[-9],Metas!C1:C26,13,0)/RC[54])*RC[-2],""Sin Metas""),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-10],'Fuente Hogares'!C1:C2,2,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-2]=0,RC[-1]>=1),100%,IFERROR(IF(RC[-2]=0,100%,RC[-1]/RC[-2]),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-12],Metas!C1:C26,12,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=ROUND(IFERROR((VLOOKUP(RC[-13],Metas!C1:C25,12,0)/RC[50])*RC[-6],""Sin Metas""),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(IFERROR(VLOOKUP(RC[-14],MOVIL!C1:C26,5,FALSE),0),IFERROR(VLOOKUP(Liquidacion_PART_TIME!RC[-14],MOVIL!C1:C26,9,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(IF(RC[-2]=0,100%,RC[-1]/RC[-2]),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=+RC[-4]+RC[-8]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=+RC[-8]+RC[-4]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]+RC[-8]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(IF(RC[-2]=0,100%,RC[-1]/RC[-2]),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-20],Metas!C[-22]:C[2],14,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR((RC[-1]/RC[42])*RC[-14],""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(IFERROR(VLOOKUP(RC[-22],MOVIL!C1:C26,11,FALSE),0),IFERROR(VLOOKUP(RC[-22],MOVIL!C1:C26,12,FALSE),0),IFERROR(VLOOKUP(Liquidacion_PART_TIME!RC[-22],'Fuente Hogares'!C7:C8,2,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-2]=0,RC[-1]>=1),100%,IF(AND(RC[-3]=0,RC[-1]=0),0%,IFERROR(IF(RC[-2]=0,100%,RC[-1]/RC[-2]),0)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(IFERROR(VLOOKUP(RC[-24],'Resumen Puntos'!C1:C33,32,FALSE),0),IFERROR(VLOOKUP(RC[-24],'Resumen Puntos'!C1:C33,30,FALSE),0),IFERROR(VLOOKUP(RC[-24],'Resumen Puntos'!C1:C55,54,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-25],'Resumen Puntos'!C1:C74,74,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[1]/(RC[-21]*RC[-19]),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-3]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-28],Metas!C1:C8,8,0),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR((VLOOKUP(RC[-29],Metas!C1:C8,8,0)/RC[34])*RC[-22],""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(SUM(IFERROR(VLOOKUP(RC[-30],'Fuente Hogares'!C26:C27,2,FALSE),0),IFERROR(VLOOKUP(Liquidacion_PART_TIME!RC[-30],'Fuente Hogares'!C29:C30,2,FALSE),0),0)) - (SUM(IFERROR(VLOOKUP(RC[-30],'Fuente Hogares'!C40:C41,2,FALSE),0),IFERROR(VLOOKUP(Liquidacion_PART_TIME!RC[-30],'Fuente Hogares'!C43:C44,2,FALSE),0)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-1]=0,OR(RC[6]=0,RC[6]="""")),0,IF(AND(RC[-2]<0,RC[-1]<0),RC[-2]/RC[-1],IF(AND(RC[-2]<0,RC[-1]>=0),140%,IF(AND(RC[-2]>0,RC[-1]>0),RC[-1]/RC[-2],IF(AND(RC[-2]>0,RC[-1]<=0),0%)))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(LOOKUP(RC[-1],'Tabla Var'!R3C13:R13C15)=""lineal"",RC[-1],LOOKUP(RC[-1],'Tabla Var'!R2C13:R13C15))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[12]=""Sin Metas"",RC[12]=""N/A"",RC[12]=0),RC[-1]*25%,RC[-1]*R2C36)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]*(RC[-27]/RC[29]))*RC[-29]*RC[29]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-35],'NPS-UMBRAL'!C1:C6,6,0),""SIN OFC"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR((VLOOKUP(RC[-36],'NPS-UMBRAL'!C1:C15,6,0)/RC[27])*RC[-29],""SIN OFC"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-37],'NPS-UMBRAL'!C1:C15,5,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-39],Metas!C1:C26,15,0),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-40],'NPS-UMBRAL'!C1:C17,15,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(IF(RC[-1]<=0%,0%,RC[-1]/RC[-2]),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-4]>=80%,IF(LOOKUP(RC[-1],'Tabla Var'!R3C1:R13C3)=""lineal"",RC[-1],LOOKUP(RC[-1],'Tabla Var'!R3C1:R13C3)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]*R2C46,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]*(RC[-37]/RC[19]))*RC[-39]*RC[19]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-45],Metas!C1:C26,19,0),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-1]=""Sin Metas"",RC[-1]=0,RC[-1]=""N/A""),0,IF(AND((ROUND(IFERROR((VLOOKUP(RC[-46],Metas!C[-48]:C[-23],19,0)/RC[17])*RC[-39],""Sin Metas""),0))>=0,(ROUND(IFERROR((VLOOKUP(RC[-46],Metas!C[-48]:C[-23],19,0)/RC[17])*RC[-39],""Sin Metas""),0))<=0.5),1,ROUND(IFERROR((VLOOKUP(RC[-46],Metas!C[-48]:C[-23],19,0)/RC[17])*RC[-39],""Sin Metas""),0)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-47],MOVIL!C1:C21,21,FALSE),0) + IFERROR(VLOOKUP(RC[-47],'Fuente Hogares'!C33:C34,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-2],0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(LOOKUP(RC[-1],'Tabla Var'!R3C13:R13C15)=""lineal"",RC[-1],LOOKUP(RC[-1],'Tabla Var'!R2C13:R13C15))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]*R2C53,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]*(RC[-44]/RC[12]))*RC[-46]*RC[12]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-19]+RC[-9]+RC[-2]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-23]=0,OR(RC[-16]=0,RC[-16]="""")),0,((RC[-1]*RC[-48]*RC[-46])*(RC[-46]/RC[10])))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-28]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-25]=0,OR(RC[-18]=0,RC[-18]="""")),0,SUM((RC[-2],RC[-28],RC[19]:RC[26])))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[14]>0,RC[-1]>0,RC[-1]<RC[14]),0,IF(AND(RC[14]>0,RC[-1]>0,RC[-1]>RC[14]),RC[-1]-RC[14],RC[-1]))"
    ActiveCell.Offset(0, 5).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-61],'Ausentismos-Vaca-Umb'!C[-63]:C[-57],7,0),0)+IFERROR(IF(VLOOKUP(RC[-61],'Ausentismos-Vaca-Umb'!C19:C29,11,0)>8,VLOOKUP(RC[-61],'Ausentismos-Vaca-Umb'!C19:C30,11,0),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-62],'Ausentismos-Vaca-Umb'!C[-54]:C[-50],5,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "240"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C9,Meses!C[-66]:C[-63],4,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-65],HC!C3:C6,4,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-1],1,4)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-2],5,2)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-3],7,2)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-69],'Ausentismos-Vaca-Umb'!C[-71]:C[-67],5,0),""S/N"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-70],Garantizado!C3:C5,3,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-71],HC!C3:C20,18,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-72],HC!C3:C12,10,0)"
    ActiveCell.Offset(0, 2).Range("A1:B1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-74],INCENTIVO!C1:C7,7,FALSE),0)"
'    ActiveCell.Offset(0, 2).Range("A1:B1").Select
'    ActiveCell.FormulaR1C1 = _
'        "=IFERROR(VLOOKUP(RC[-76],'INCENTIVO MOTOROLA - CLARO'!C1:C2,2,FALSE),0)"
'    ActiveCell.Offset(0, 2).Range("A1:B1").Select
'    ActiveCell.FormulaR1C1 = _
'        "=IFERROR(VLOOKUP(RC[-78],'INCENTIVO MOTOROLA - CLARO'!C1:C3,3,FALSE),0)"
'    ActiveCell.Offset(0, 2).Range("A1:B1").Select
'    ActiveCell.FormulaR1C1 = _
'        "=IFERROR(VLOOKUP(RC[-80],'INCENTIVO POSPAGO'!C1:C7,7,FALSE),0)"
    ActiveCell.Offset(1, 0).Range("A1:B1").Select
    Selection.End(xlToLeft).Select
    
    Next i
    


 Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
marcador:
Application.Calculation = xlCalculationAutomatic
MsgBox "Listo hoja Part-time, final final", vbInformation
MsgBox "proceso FINAL FINAL completado , puede continuar con el siguiente paso", vbInformation
MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA ABRIR EL ARCHIVO Y EJECUTE EL BOTÓN CÁLCULOS AUXILIARES", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.Quit

End Sub

Sub tiendas()


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Select
Range("I2") = UserForm7.ComboBox1
Range("a1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C[2])-2"
vari_9 = Range("a1").Value
Range("c4").Select

For i = 1 To vari_9

  ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],HC!C3:C5,3,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],HC!C3:C7,5,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-3],Metas!C1:C5,3,0),""Sin Oficina"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-4],Metas!C1:C5,5,0),""Sin Oficina"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(VLOOKUP(RC[-5],HC!C3:C9,7,0),0)>0,VLOOKUP(RC[-5],'DESARROLLO+PROYECTOS'!C3:C9,7,FALSE),IFERROR(VLOOKUP(RC[-5],HC!C3:C9,7,0),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C9,Meses!C[-8]:C[-7],2,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
     ActiveCell.FormulaR1C1 = _
        "=IF((IF((IF(AND(RC[56]=""retiro"",RC[53]=RC[51],RC[54]=RC[-1]),IF((IF(RC[52]>(CONCATENATE(RC[51],RC[-1],RC[50])-RC[48]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[50]-RC[55])-RC[48])+1,(RC[50]-RC[55]+1)+(RC[50]-(RC[48]+RC[49]))),IF(AND(RC[53]=RC[51],RC[54]=RC[-1]),(RC[50]-RC[55]+1)-(RC[48]+RC[49]),RC[50]-(RC[48]+RC[49]))))<0,0,IF(AND(RC[56]=""retiro"",RC[53]=RC[51],RC[54]=RC[-1]),IF((IF(RC[52]>(CONCATENATE(RC[51],RC[-1],RC[50])-RC[48]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[50]-RC[55])-RC[48])+1,(RC[50]-RC[55]+1)+(RC[50]-(RC[48]+RC[49]))),IF(AND(RC[53]=RC[51],RC[54]=RC[-1]),(RC[50]-RC[55]+1)-(RC[48]+RC[49]),RC[50]-(RC[48]+RC[49])))-IFERROR(VLOOKUP(RC[-7],'AISLAMIENTO COVID'!C1:C7,7,FALSE),0)))<0,0," & _
        "IF((IF(AND(RC[56]=""retiro"",RC[53]=RC[51],RC[54]=RC[-1]),IF((IF(RC[52]>(CONCATENATE(RC[51],RC[-1],RC[50])-RC[48]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[50]-RC[55])-RC[48])+1,(RC[50]-RC[55]+1)+(RC[50]-(RC[48]+RC[49]))),IF(AND(RC[53]=RC[51],RC[54]=RC[-1]),(RC[50]-RC[55]+1)-(RC[48]+RC[49]),RC[50]-(RC[48]+RC[49]))))<0,0,IF(AND(RC[56]=""retiro"",RC[53]=RC[51],RC[54]=RC[-1]),IF((IF(RC[52]>(CONCATENATE(RC[51],RC[-1],RC[50])-RC[48]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[50]-RC[55])-RC[48])+1,(RC[50]-RC[55]+1)+(RC[50]-(RC[48]+RC[49]))),IF(AND(RC[53]=RC[51],RC[54]=RC[-1]),(RC[50]-RC[55]+1)-(RC[48]+RC[49]),RC[50]-(RC[48]+RC[49])))-IFERROR(VLOOKUP(RC[-7],'AISLAMIENTO COVID'!C1:C7,7,FALSE),0)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-8],Metas!C1:C26,13,0),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-2]<=0,0,IF(OR(RC[-1]=0,RC[-1]=""Sin Metas""),0,IF((ROUND(IFERROR(VLOOKUP(RC[-9],Metas!C1:C25,13,0)*(RC[-2]/RC[48]),""Sin Metas""),0))<=0,1,(ROUND(IFERROR(VLOOKUP(RC[-9],Metas!C1:C25,13,0)*(RC[-2]/RC[48]),""Sin Metas""),0)))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(IFERROR(VLOOKUP(RC[-10],'Fuente Hogares'!C1:C2,2,0),0),IFERROR(VLOOKUP(RC[-10],'Fuente Hogares'!C78:C79,2,0),0),IFERROR(VLOOKUP(RC[-10],'Fuente Hogares'!C82:C83,2,0),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-4]=0,0,IF(AND(RC[-2]=0,RC[-1]>=1),100%,IFERROR(IF(RC[-2]=0,100%,RC[-1]/RC[-2]),0)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-12],Metas!C1:C26,12,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-6]<=0,0,IF(OR(RC[-1]=0,RC[-1]=""Sin Metas""),0,IF((ROUND(IFERROR(VLOOKUP(RC[-13],Metas!C1:C25,12,0)*(RC[-6]/RC[44]),""Sin Metas""),0))<=0,1,(ROUND(IFERROR(VLOOKUP(RC[-13],Metas!C1:C25,12,0)*(RC[-6]/RC[44]),""Sin Metas""),0)))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(IFERROR(VLOOKUP(RC[-14],MOVIL!C1:C10,10,FALSE),0),IFERROR(VLOOKUP(RC[-14],MOVIL!C1:C14,14,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-8]=0,0,IFERROR(IF(RC[-2]=0,100%,RC[-1]/RC[-2]),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(RC[-2],IF(AND(RC[-1]>=90%,RC[5]>=80%),SUM(VLOOKUP(RC[-16],Resumen_plan_power!C1:C5,5,FALSE),VLOOKUP(RC[-16],Resumen_plan_power!C1:C7,7,FALSE),VLOOKUP(RC[-16],Resumen_plan_power!C1:C17,17,FALSE)),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-10]=0,0,IFERROR(IF(RC[-4]=0,100%,RC[-1]/RC[-4]),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=+RC[-6]+RC[-10]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=+RC[-10]+RC[-6]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-6]+RC[-10]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-14]=0,0,IFERROR(IF(RC[-2]=0,100%,RC[-1]/RC[-2]),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-6]+RC[-12]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-16]=0,0,IFERROR(IF(AND(RC[-4]=0,RC[-1]=0),0%,RC[-1]/RC[-4]),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-24],Metas!C[-26]:C[-2],14,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-18]<=0,0,IFERROR(RC[-1]*(RC[-18]/RC[32]),""Sin Metas""))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(IFERROR(VLOOKUP(RC[-26],MOVIL!C1:C15,15,FALSE),0),IFERROR(VLOOKUP(RC[-26],MOVIL!C1:C16,16,FALSE),0),IFERROR(VLOOKUP(RC[-26],'Fuente Hogares'!C7:C8,2,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-20]=0,0,IF(AND(RC[-2]=0,RC[-1]>=1,),100%,IF(AND(RC[-3]=0,RC[-1]=0),0%,IFERROR(IF(RC[-2]=0,100%,RC[-1]/RC[-2]),0))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-28],'Resumen Puntos'!C1:C100,100,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-29],'Resumen Puntos'!C1:C101,101,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-25]=0,RC[-7]<80%),0,IF(AND(RC[-25]=0,RC[-7]>160%),160%*80%,IF(RC[-25]=0,RC[-7]*80%,IFERROR(RC[-2]/RC[-25],0))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-26]=0,0,RC[-2]+RC[-3])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-32],Metas!C1:C26,15,0),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(VLOOKUP(RC[-33],'NPS-UMBRAL'!C1:C17,15,0),0)=""NO REGISTRA"",""NO REGISTRA ENCUESTAS"",IFERROR(VLOOKUP(RC[-33],'NPS-UMBRAL'!C1:C17,15,0),""NO REGISTRA ENCUESTAS""))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-27]=0,0%,IFERROR(IF(RC[-1]<=0%,0%,RC[-1]/RC[-2]),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[36]>=100%,IF(RC[-2]=100%,160%,IF(VLOOKUP(RC[-1],'Tabla Var'!R2C5:R13C7,3,TRUE)=""lineal"",'Liquid Tiendas-Cvc-DENTRO CAV'!RC[-1],VLOOKUP(RC[-1],'Tabla Var'!R2C5:R13C7,3,TRUE))),IF(VLOOKUP(RC[-1],'Tabla Var'!R2C5:R13C8,4,TRUE)=""lineal"",'Liquid Tiendas-Cvc-DENTRO CAV'!RC[-1],VLOOKUP(RC[-1],'Tabla Var'!R2C5:R13C8,4,TRUE)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[2]=""Sin Metas"",RC[2]=""N/A"",RC[2]=0),RC[-1]*20%,RC[-1]*R2C39)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(RC[-1]*((SUM(RC[-30],IF(IFERROR(VLOOKUP(RC[-37],'NPS-UMBRAL'!C1:C10,10,FALSE),0)>8,IFERROR(VLOOKUP(RC[-37],'NPS-UMBRAL'!C1:C10,10,FALSE),0),0))/RC[20]))*RC[-32])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-38],Metas!C1:C26,19,0),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-32]<=0,0,IF(OR(RC[-1]=""Sin Metas"",RC[-1]=0,RC[-1]=""N/A""),0,IF(AND((ROUND(IFERROR(VLOOKUP(RC[-39],Metas!C[-41]:C[-16],19,0)*(RC[-32]/RC[18]),""Sin Metas""),0))>=0,(ROUND(IFERROR(VLOOKUP(RC[-39],Metas!C[-41]:C[-16],19,0)*(RC[-32]/RC[18]),""Sin Metas""),0))<=0.5),1,ROUND(IFERROR(VLOOKUP(RC[-39],Metas!C[-41]:C[-16],19,0)*(RC[-32]/RC[18]),""Sin Metas""),0))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-40],MOVIL!C1:C25,25,FALSE),0) + IFERROR(VLOOKUP(RC[-40],'Fuente Hogares'!C33:C34,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-34]=0,0,IFERROR(RC[-1]/RC[-2],0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(LOOKUP(RC[-1],'Tabla Var'!R3C13:R13C15)=""lineal"",RC[-1],LOOKUP(RC[-1],'Tabla Var'!R2C13:R13C15))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]*R2C46,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(RC[-1]*((SUM(RC[-37],IF(IFERROR(VLOOKUP(RC[-44],'NPS-UMBRAL'!C1:C10,10,FALSE),0)>8,IFERROR(VLOOKUP(RC[-44],'NPS-UMBRAL'!C1:C10,10,FALSE),0),0))/RC[13]))*RC[-39])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-9]+RC[-2]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=((RC[-1]*RC[-41])*((SUM(RC[-39],IF(IFERROR(VLOOKUP(RC[-46],'NPS-UMBRAL'!C1:C10,10,FALSE),0)>8,IFERROR(VLOOKUP(RC[-46],'NPS-UMBRAL'!C1:C10,10,FALSE),0),0))/RC[11])))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-17]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-43]=0,RC[-41]=0),0,VLOOKUP(RC[-48],Resumen_plan_power!C1:C15,15,FALSE))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3],RC[-18],RC[-1],RC[18])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[14]>0,RC[-1]>0,RC[-1]<RC[14]),0,IF(AND(RC[14]>0,RC[-1]>0,RC[-1]>RC[14]),RC[-1]-RC[14],RC[-1]))"
    ActiveCell.Offset(0, 5).Range("A1").Select
ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-55],'Ausentismos-Vaca-Umb'!C[-57]:C[-51],7,0),0)+IFERROR(IF(VLOOKUP(RC[-55],'NPS-UMBRAL'!C1:C11,10,0)>8,VLOOKUP(RC[-55],'NPS-UMBRAL'!C1:C11,10,0),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-56],'Ausentismos-Vaca-Umb'!C[-48]:C[-44],5,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C9,Meses!C[-59]:C[-57],3,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C9,Meses!C[-60]:C[-57],4,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-59],HC!C3:C6,4,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-1],1,4)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-2],5,2)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-3],7,2)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-63],'Ausentismos-Vaca-Umb'!C[-65]:C[-61],5,0),""S/N"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-64],Garantizado!C3:C5,3,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-65],HC!C3:C20,18,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-66],HC!C3:C12,10,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-60]=0,0,IFERROR(VLOOKUP(RC[-67],'DESARROLLO+PROYECTOS'!C3:C34,32,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-27],RC[-34],RC[-41],RC[-47])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-69],Metas!C41:C46,6,0),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "='Tabla Var'!R16C2"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-2]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-72],Metas!C41:C48,8,FALSE),""Sin metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-66]<=0,0,IF(OR(RC[-1]=0,RC[-1]=""Sin Metas""),0,IF((ROUND(IFERROR(VLOOKUP(RC[-73],Metas!C41:C48,8,0)*(RC[-66]/RC[-16]),""Sin Metas""),0))<=0,1,(ROUND(IFERROR(VLOOKUP(RC[-73],Metas!C41:C48,8,0)*(RC[-66]/RC[-16]),""Sin Metas""),0)))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-74],'Fuente Hogares'!C87:C88,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]=0,100%,IFERROR(RC[-1]/RC[-2],0))"
    ActiveCell.Offset(1, 0).Select
    Selection.End(xlToLeft).Select


    
 Next i
 
 
 Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation



End Sub

Sub proceso_1()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim archivos

Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\FIJA\*.xlsb")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\FIJA\" & archivos
archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\FIJAS SOLO CAV")


xiaomi = "FIJAS SOLO CAV"
    ss = xiaomi & ".xlsx"

Windows(ss).Activate


Windows(tt).Activate
Sheets("BVL").Select
    Set rangodatos = Sheets("BVL").UsedRange
    rangodatos.AutoFilter Field:=25, Criteria1:="1"

ultima = Sheets("BVL").Range("A" & Rows.Count).End(xlUp).Row

   Sheets("BVL").Range("A1:CQ" & ultima).Copy

    Windows(ss).Activate

Sheets("Hoja1").Select
Range("A2").Select
ActiveSheet.Paste

Rows("2:2").EntireRow.Delete
ActiveWorkbook.Save
ActiveWorkbook.Close

Windows(tt).Activate
Application.DisplayAlerts = False
ActiveWorkbook.Close
Application.DisplayAlerts = True

Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "proceso FINAL FINAL completado , puede continuar con el siguiente paso", vbInformation


End Sub


Sub AUSENT()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos

Dim vari_2 As Excel.Workbook
Dim vari_3 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\AUSENTISMOS\*.xlsx")
Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\AUSENTISMOS\" & archivos
archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate

Sheets(1).Select
Range("A1").Select
ActiveSheet.ShowAllData

col = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("A2:H" & col).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Ausentismos-Vaca-Umb").Select
Range("A2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False



'+++
archii = Dir("D:\AUTOMATIZACION\PARTIME\*.xls")
Do While archii <> ""
Workbooks.Open "D:\AUTOMATIZACION\PARTIME\" & archii
archii = Dir
Loop

vari_4 = ActiveWorkbook.Name

pp = vari_4

Windows(pp).Activate
On Error GoTo marcador
Sheets("Comisiones").Select
Range("A1").Select


peter = Sheets("Comisiones").Range("A" & Rows.Count).End(xlUp).Row
Sheets("Comisiones").Range("A2:F" & peter).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Ausentismos-Vaca-Umb").Select
Range("AE2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

Windows(pp).Activate
ActiveWorkbook.Close SaveChanges:=False




Application.ScreenUpdating = True
Application.DisplayAlerts = True

marcador:
MsgBox "proceso FINAL FINAL completado , puede continuar con el siguiente paso", vbInformation


End Sub


Sub CARGUE_HC()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim vari_2 As Excel.Workbook
Dim vari_3 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\PLANTILLA_HC")
xiaomi = "PLANTILLA_HC"
    ss = xiaomi & ".xlsx"

Windows(ss).Activate
col = Sheets("HC").Range("A" & Rows.Count).End(xlUp).Row
Sheets("HC").Range("A2:U" & col).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("HC").Select
Range("C2").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=False



Range("AC1").Select
 ActiveCell.FormulaR1C1 = "=COUNTA(C[-26])-1"
vari_5 = Range("AC1").Value
Range("A2").Select

For i = 1 To vari_5

 ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[2],Metas!C1:C5,5,0)"

ActiveCell.Offset(1, 0).Select

Next i

Range("B2").Select

For i = 1 To vari_5

ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],Metas!C1:C26,18,0)"

ActiveCell.Offset(1, 0).Select

Next i


Range("Y2").Select

For i = 1 To vari_5

ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-22],'CEDULAS PILOTO'!C1,1,FALSE)"

ActiveCell.Offset(1, 0).Select

Next i
Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "proceso FINAL FINAL completado , puede continuar con el siguiente paso", vbInformation
MsgBox "EL EJECUTABLE DE EXCEL SE CERRARÁ POR FAVOR ÁBRALO NUEVAMENTE Y EJECUTE EL BOTÓN DE CARGUE-AUSENTISMOS", vbInformation


Windows("PLAN_LIQ.xlsm").Activate
ActiveWorkbook.Close SaveChanges:=True
Application.Quit

End Sub

Sub INCENTIVOS_pervariable_5()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim vari_2 As Excel.Workbook
Dim vari_3 As Excel.Workbook
Dim vari_6 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Dim archivos As String


archivos = Dir("D:\AUTOMATIZACION\INCENTIVOS\*.xlsx")
If archivos = "" Then
MsgBox "NO HAY NINGÚN ARCHIVO DE INCENTIVOS CARGADO", vbCritical
Else
Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\INCENTIVOS\" & archivos
archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1


Windows(tt).Activate

Sheets("Incentivo").Select
col = Sheets("Incentivo").Range("B" & Rows.Count).End(xlUp).Row
Sheets("Incentivo").Range("b4:h" & col).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("INCENTIVO").Select
Range("A1").Select
 Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False

Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "proceso FINAL FINAL completado , puede continuar con el siguiente paso", vbInformation
End If


End Sub


Sub ced_tiendas()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLAN_LIQ.xlsm").Activate
Sheets("HC").Select
Range("C1").Select
vari_7 = Selection.End(xlDown).Row
Set rangodatos = Sheets("HC").Range("A1:AB" & vari_7)
rangodatos.AutoFilter Field:=7, Criteria1:=Array("Consultor Servicio Personalizado A Clientes", "Consultor Integral Servicio A Clientes", "Asesor Servicio Al Cliente", "Consultor Integral Servicio A Clientes Sr", "Asesor Integral Servicio Al Cliente", "Consultor(a) Integral Servicio A Clientes", "Asesor(a) Servicio Al Cliente", "Consultor(a) Integral Servicio A Clientes Sr", "Consultor(a) Servicio Personalizado A Clientes", "Asesor(a) Integral Servicio Al Cliente"), _
        Operator:=xlFilterValues
Range("C1").Select
vari_7 = Selection.End(xlDown).Row
Range("AB1") = "*"
Range("AB1").Select
ActiveCell.Copy
Range("AB1:AB" & vari_7).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("AB1").Select
ActiveSheet.ShowAllData
Selection.AutoFilter
rangodatos.AutoFilter Field:=28, Criteria1:="*"
rangodatos.AutoFilter Field:=23, Criteria1:="<>Operacion Cavs", Operator:=xlAnd
Range("C1").Select
vari_7 = Selection.End(xlDown).Row
Range("AA1") = "NO SE LIQUIDA"
Range("AA1").Select
ActiveCell.Copy
Range("AA1:AA" & vari_7).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("AA1").Select
ActiveSheet.ShowAllData
Range("AA1") = "OBSERVACIONES"
Range("AA1").Select
Selection.Font.Bold = True
With Selection.Font
    .Color = -16776961
    .TintAndShade = 0
End With
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent5
    .TintAndShade = 0.799981688894314
    .PatternTintAndShade = 0
End With
MsgBox "Listo primera fase", vbInformation
rangodatos.AutoFilter Field:=28, Criteria1:="*"
rangodatos.AutoFilter Field:=27, Criteria1:="="
Range("C1").Select
vari_7 = Selection.End(xlDown).Row
Range("Z1").Select
ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-23],'ROLES_EXPERIENCIA AL CLIENTE'!C[-25]:C[-24],2,FALSE)"
ActiveCell.Copy
Range("Z1:Z" & vari_7).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("Z1") = "ROLES"
Range("Z1").Select
Selection.Font.Bold = True
With Selection.Font
    .Color = -16776961
    .TintAndShade = 0
End With
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 65535
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With
Range("Z1").Select
ActiveSheet.ShowAllData
Selection.AutoFilter

MsgBox "Listo segunda fase", vbInformation
MsgBox "proceso FINAL FINAL completado , puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True


End Sub


Sub demasbases()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
On Error Resume Next

Dim archivos

Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String


archivos = Dir("D:\AUTOMATIZACION\CAMBIOS\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\CAMBIOS\" & archivos

archivos = Dir
Loop
'
vari_1 = ActiveWorkbook.Name
'
tt = vari_1


'Windows(tt).Activate
'Sheets(1).Select
'    Set RangoDatos = Sheets(1).UsedRange
'    RangoDatos.AutoFilter Field:=18, Criteria1:="<>"
'
'Range("a1").Select
'Selection.End(xlToRight).Select
'ActiveCell.Offset(0, 1).Select
'ActiveCell.Offset(0, 0) = "CONSULTOR"
'ActiveCell.Offset(0, 0).Select
'Selection.Copy
'ActiveCell.Offset(0, -1).Select
'Selection.End(xlDown).Select
'ActiveCell.Offset(0, 1).Select
'Range(Selection, Selection.End(xlUp)).Select
'ActiveSheet.Paste
' Application.CutCopyMode = False
' Range("A1").Select
' ActiveSheet.ShowAllData


Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)
'
Worksheets.Add
ActiveSheet.Name = "Hoja de pivote"

Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("A1"))

'   With ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("CONSULTOR")
'        .Orientation = xlPageField
'        .Position = 1
'    End With
'    With ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("CEDULA")
'        .Orientation = xlRowField
'        .Position = 1
'    End With
'    ActiveSheet.PivotTables("Tabla dinámica12").AddDataField ActiveSheet. _
'        PivotTables("Tabla dinámica12").PivotFields("VLR_CFM_NETO"), _
'        "Suma de VLR_CFM_NETO", xlSum
'    With ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("GRADE")
'        .Orientation = xlPageField
'        .Position = 1
'    End With
'    ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("CONSULTOR"). _
'        ClearAllFilters
'    ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("CONSULTOR"). _
'        CurrentPage = "(blank)"
'    ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("GRADE").CurrentPage = _
'        "(All)"
'    With ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("GRADE")
'        .PivotItems("PG").Visible = False
'    End With
'    ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("GRADE"). _
'        EnableMultiplePageItems = True


  With ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("CEDULA")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica12").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica12").PivotFields("VL_CFM_NETO"), "Suma de VL_CFM_NETO", xlSum
    With ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("Movil_Fija")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("Movil_Fija"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("Movil_Fija"). _
        CurrentPage = "Movil"
        
ActiveSheet.PivotTables("Tabla dinámica12").PivotSelect "", xlDataAndLabel, True
Selection.Copy
Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("AC9").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
 Application.CutCopyMode = False
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False


'--
archi = Dir("D:\AUTOMATIZACION\FIDELIZACION\*.xlsx")

Do While archi <> ""
Workbooks.Open "D:\AUTOMATIZACION\FIDELIZACION\" & archi

archi = Dir
Loop
'
tigre = ActiveWorkbook.Name
'
pp = tigre

Windows(pp).Activate


Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)
'
Worksheets.Add
ActiveSheet.Name = "Hoja de pivote"

Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("A1"))

With ActiveSheet.PivotTables("Tabla dinámica13").PivotFields("CEDULA")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica13").AddDataField ActiveSheet. _
        PivotTables("Tabla dinámica13").PivotFields("CFM"), "Suma de CFM", xlSum


ActiveSheet.PivotTables("Tabla dinámica13").PivotSelect "", xlDataAndLabel, True
Selection.Copy


Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("AN7").Select
ActiveSheet.Paste
 Application.CutCopyMode = False
Windows(pp).Activate
ActiveWorkbook.Close SaveChanges:=False




'--
arch = Dir("D:\AUTOMATIZACION\NO_RETENIDOS\*.xlsx")

Do While arch <> ""
Workbooks.Open "D:\AUTOMATIZACION\NO_RETENIDOS\" & arch

arch = Dir
Loop
'
vari_8 = ActiveWorkbook.Name
'
mm = vari_8

Windows(mm).Activate

Sheets(1).Select

 Columns("K:K").Select
    Selection.TextToColumns Destination:=Range("K1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Range("K1").Select

Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets(1).Range("A1").CurrentRegion.Address)
'
Worksheets.Add
ActiveSheet.Name = "Hoja de pivote"

Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("A1"))



With ActiveSheet.PivotTables("Tabla dinámica15").PivotFields("Cedula ASESOR")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica15").AddDataField ActiveSheet. _
        PivotTables("Tabla dinámica15").PivotFields("RENTA_CAN"), "Suma de RENTA_CAN", _
        xlSum
        
        ActiveSheet.PivotTables("Tabla dinámica15").PivotSelect "", xlDataAndLabel, True
Selection.Copy


Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("AQ7").Select
ActiveSheet.Paste
 Application.CutCopyMode = False
Windows(mm).Activate
ActiveWorkbook.Close SaveChanges:=False
Range("A1").Select

Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "proceso FINAL FINAL completado , puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation

End Sub


Sub FORMATITOS_1()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("Liquidacion").Select
'If Range("A1") > 1 Then

'Sheets("Liquidacion").Select
'Range("F4").Select
'ActiveCell.Range("A1:B1").Select
'    With Selection
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
'    End With
'    Selection.Merge
'    Selection.Copy
'    ActiveCell.Offset(1, 0).Range("A1:B1").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False
'    ActiveCell.Offset(-2, 0).Range("A1:B1").Select
'
'Range("J4").Select
'
'Else: Range("F4:G4").Select
'    With Selection
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
'    End With
'    Selection.Merge
'    Range("C3").Select
'End If
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
'MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub
'***************
Sub FORMATITOS_2()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

Sheets("Liquidacion_PART_TIME").Select
If Range("A1") > 1 Then

Sheets("Liquidacion_PART_TIME").Select
Range("F4").Select
ActiveCell.Range("A1:B1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1:B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(-2, 0).Range("A1:B1").Select

Range("J4").Select

Else: Range("F4:G4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C3").Select
End If
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
End Sub
'***************
Sub FORMATITOS_3()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Select


'If Range("A1") > 1 Then
'
'Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Select
'Range("F4").Select
'ActiveCell.Range("A1:B1").Select
'    With Selection
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
'    End With
'    Selection.Merge
'    Selection.Copy
'    ActiveCell.Offset(1, 0).Range("A1:B1").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
'        SkipBlanks:=False, Transpose:=False
'    ActiveCell.Offset(-2, 0).Range("A1:B1").Select
'
'Range("J4").Select
'
'Else: Range("F4:G4").Select
'    With Selection
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
'    End With
'    Selection.Merge
'    Range("C3").Select
'End If
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
'MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
End Sub
Sub FORMATITOS_4()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("Liquidacion").Select
Range("a1").Select
vari_9 = Range("a1").Value
Range("A4").Select
For i = 1 To vari_9

 ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[2],HC!C3:C13,11,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],HC!C3:C4,2,0)"
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlToLeft).Select
Next i

MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
End Sub
Sub FORMATITOS_5()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("Liquidacion_PART_TIME").Select
Range("a1").Select
vari_9 = Range("a1").Value
Range("A4").Select
For i = 1 To vari_9

On Error GoTo marcador
 ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[2],HC!C3:C13,11,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],HC!C3:C4,2,0)"
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlToLeft).Select
Next i
marcador:
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
End Sub
Sub FORMATITOS_6()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Select
Range("a1").Select
vari_9 = Range("a1").Value
Range("A4").Select
For i = 1 To vari_9

 ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[2],HC!C3:C13,11,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],HC!C3:C4,2,0)"
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlToLeft).Select
Next i

MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic






End Sub
Sub FORMATITOS_8()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("OTRAS GESTIONES").Select
Range("a1").Select
vari_9 = Range("a1").Value
Range("A4").Select
For i = 1 To vari_9

 ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[2],HC!C3:C13,11,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],HC!C3:C4,2,0)"
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlToLeft).Select
Next i

MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic






End Sub


Sub guarda()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
  
  
  Dim archivos

Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

Call booo
  
Sheets("Meses").Select
Range("M1") = UserForm6.ComboBox1
Range("M2") = UserForm6.ComboBox2
  
  
  RUTA = "D:\AUTOMATIZACION\LIQ_FINAL\"
    GENERAR_ARCHIVO_EXCEL = RUTA & "Liq_" & Range("M1").Value & Range("M2").Value & "_Consultor Integral Servicio al Cliente" & ".xlsx"
   
    ActiveWorkbook.SaveAs Filename:= _
        GENERAR_ARCHIVO_EXCEL, FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False

vari_1 = ActiveWorkbook.Name

tt = vari_1




Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLAN_LIQ")



xiaomi = "PLAN_LIQ"
    ss = xiaomi & ".xlsm"





Windows(ss).Activate



Windows(tt).Activate
Call vari_10



Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub


Sub mapeotop()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
On Error Resume Next

Dim archivos

Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String


archivos = Dir("D:\AUTOMATIZACION\MAPEO\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\MAPEO\" & archivos

archivos = Dir
Loop
'
vari_1 = ActiveWorkbook.Name
'
tt = vari_1


Windows(tt).Activate
Sheets(1).Select


 Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Range("A1").Select
    
    
    
Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets(1).Range("A1").CurrentRegion.Address)
'
Worksheets.Add
ActiveSheet.Name = "Hoja de pivote"

Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("A1"))


With ActiveSheet.PivotTables("Tabla dinámica17").PivotFields("CEDULA")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("UND_PORTA_POS_MASIVO_CONSULTOR_SIN_APP"), _
        "Suma de UND_PORTA_POS_MASIVO_CONSULTOR_SIN_APP", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("UND_POS_MASIVO_CONSULTOR_SIN_APP"), _
        "Suma de UND_POS_MASIVO_CONSULTOR_SIN_APP", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("UND_MIGRA_MASIVO_CONSULTOR_SIN_APP"), _
        "Suma de UND_MIGRA_MASIVO_CONSULTOR_SIN_APP", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("TOTAL _UND_POSPAGO_MASIVO_CONSULTOR_SIN_APP"), _
        "Suma de TOTAL _UND_POSPAGO_MASIVO_CONSULTOR_SIN_APP", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("UND_PORTA_POS_MASIVO_CONSULTOR_APP"), _
        "Suma de UND_PORTA_POS_MASIVO_CONSULTOR_APP", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("UND_POS_MASIVO_CONSULTOR_APP"), _
        "Suma de UND_POS_MASIVO_CONSULTOR_APP", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("UND_MIGRA_MASIVO_CONSULTOR_APP"), _
        "Suma de UND_MIGRA_MASIVO_CONSULTOR_APP", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("TOTAL _UND_POSPAGO_MASIVO_CONSULTOR_APP"), _
        "Suma de TOTAL _UND_POSPAGO_MASIVO_CONSULTOR_APP", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("TOTAL _UND_POSPAGO_MASIVO_CONSULTOR"), _
        "Suma de TOTAL _UND_POSPAGO_MASIVO_CONSULTOR", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("UND_PORTA_PYME_CONSULTOR"), _
        "Suma de UND_PORTA_PYME_CONSULTOR", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("UND_POS_PYME_CONSULTOR"), _
        "Suma de UND_POS_PYME_CONSULTOR", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("UND_MIGRA_PYME_CONSULTOR"), _
        "Suma de UND_MIGRA_PYME_CONSULTOR", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("TOTAL _UND_POSPAGO_PYME_CONSULTOR"), _
        "Suma de TOTAL _UND_POSPAGO_PYME_CONSULTOR", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("TOTAL_INGRESOS_CONSULTOR_TERMINALES"), _
        "Suma de TOTAL_INGRESOS_CONSULTOR_TERMINALES", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("TOTAL_INGRESOS_CONSULTOR_REPO"), _
        "Suma de TOTAL_INGRESOS_CONSULTOR_REPO", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("UND_CLARO_UP_CONSULTOR"), _
        "Suma de UND_CLARO_UP_CONSULTOR", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("UND_RETO_MOVI_CONSULTOR"), _
        "Suma de UND_RETO_MOVI_CONSULTOR", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("UND_RANGO_CFM1_CONSULTOR"), _
        "Suma de UND_RANGO_CFM1_CONSULTOR", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("UND_RANGO_CFM2_CONSULTOR"), _
        "Suma de UND_RANGO_CFM2_CONSULTOR", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields( _
        "UND_RANGO_CFM3_CONSULTOR"), _
        "Suma de UND_RANGO_CFM3_CONSULTOR", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("UND_CLARO_APP_CONSULTOR"), _
        "Suma de UND_CLARO_APP_CONSULTOR", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("UND_PAQUETE_R1_CONSULTOR"), _
        "Suma de UND_PAQUETE_R1_CONSULTOR", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("UND_PAQUETE_R2_CONSULTOR"), _
        "Suma de UND_PAQUETE_R2_CONSULTOR", xlSum
    ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("Total Clientes Convergentes Móvil Consultor Cav"), _
        "Suma de Total Clientes Convergentes Móvil Consultor Cav", xlSum
       ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("Roaming "), _
        "Suma de Roaming ", xlSum
        ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("INGRESOS_TERM_CONSULTOR"), _
        "Suma de INGRESOS_TERM_CONSULTOR", xlSum
        ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("INGRESOS_TERM_CONSULTOR_FINAN"), _
        "Suma de INGRESOS_TERM_CONSULTOR_FINAN", xlSum
        ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("INGRESOS_REPO_CONSULTOR"), _
        "Suma de INGRESOS_REPO_CONSULTOR", xlSum
        ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("INGRESOS_REPO_CONSULTOR_FINAN"), _
        "Suma de INGRESOS_REPO_CONSULTOR_FINAN", xlSum
        ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("Total_Power_Consultor_Rango_1"), _
        "Suma de Total_Power_Consultor_Rango_1", xlSum
        ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("Total_Power_Consultor_Rango_3"), _
        "Suma de Total_Power_Consultor_Rango_3", xlSum
        ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("Cavariab_7l_ 223_FO"), _
        "Suma de Cavariab_7l_ 223_FO", xlSum
        ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("CFM_CONSULTOR_<25000_SIN"), _
        "Suma de CFM_CONSULTOR_<25000_SIN", xlSum
        ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("CFM_CONSULTOR_<25000"), _
        "Suma de CFM_CONSULTOR_<25000", xlSum
        ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("Ingre_Consultor_Cfm_pyme_cav"), _
        "Suma de Ingre_Consultor_Cfm_pyme_cav", xlSum
        ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("Total_Power_Consultor_Rango_2"), _
        "Suma de Total_Power_Consultor_Rango_2", xlSum
        ActiveSheet.PivotTables("Tabla dinámica17").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica17").PivotFields("Ventas_Cloud"), _
        "Suma de Ventas_Cloud", xlSum
        
        
        
        
ActiveSheet.PivotTables("Tabla dinámica17").ShowValuesRow = False
  With ActiveSheet.PivotTables("Tabla dinámica17").PivotFields("CANAL")
        .Orientation = xlPageField
        .Position = 1
    End With
      ActiveSheet.PivotTables("Tabla dinámica17").PivotFields("CANAL"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica17").PivotFields("CANAL").CurrentPage = _
        "CAV"

    ActiveSheet.PivotTables("Tabla dinámica17").PivotSelect "", xlDataAndLabel, True
Selection.Copy

Windows("PLAN_LIQ.xlsm").Activate

Sheets("MOVIL").Select
Range("A3").Select
ActiveSheet.Paste
 Application.CutCopyMode = False
 Range("A3").Select
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False

Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "proceso FINAL FINAL completado , puede continuar con el siguiente paso", vbInformation

End Sub


Sub vari_11()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos

Dim vari_2 As Excel.Workbook

Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\METAS\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\METAS\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate

Sheets(1).Select
vegas = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("B1:D" & vegas).Copy
    
    
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("A3").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Windows(tt).Activate

Sheets(1).Select
variab_1 = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("E1:E" & variab_1).Copy
    
    
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("E3").Select
ActiveSheet.Paste
Application.CutCopyMode = False
    
    
Windows(tt).Activate

Sheets(1).Select
variab_2 = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("F1:F" & variab_2).Copy
    
    
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("G3").Select
ActiveSheet.Paste
Application.CutCopyMode = False
    
    
Windows(tt).Activate

Sheets(1).Select
variab_3 = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("M1:M" & variab_3).Copy
    
    
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("H3").Select
ActiveSheet.Paste
Application.CutCopyMode = False


Windows(tt).Activate

Sheets(1).Select
variab_4 = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("G1:I" & variab_4).Copy
    
    
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("I3").Select
ActiveSheet.Paste
Application.CutCopyMode = False


Windows(tt).Activate

Sheets(1).Select
variab_5 = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("J1:J" & variab_5).Copy
    
    
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("M3").Select
ActiveSheet.Paste
Application.CutCopyMode = False


Windows(tt).Activate

Sheets(1).Select
variab_6 = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("K1:K" & variab_6).Copy
    
    
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("N3").Select
ActiveSheet.Paste
Application.CutCopyMode = False


Windows(tt).Activate

Sheets(1).Select
variab_7 = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("L1:L" & variab_7).Copy
    
    
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("S3").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Windows(tt).Activate

Sheets(1).Select
ny = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("O1:P" & ny).Copy
    
    
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("T3").Select
ActiveSheet.Paste
Application.CutCopyMode = False



Windows(tt).Activate

Sheets(1).Select
variab_8 = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("N1:N" & variab_8).Copy
    
    
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("O3").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Rows(3).EntireRow.Delete

Range("A2").Select
vari_7 = Selection.End(xlDown).Row

Range("L3").Select

ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
ActiveCell.Copy

Range("L4:L" & vari_7).Select
ActiveSheet.Paste
Application.CutCopyMode = False


Range("R3").Select

ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-13],'Tabla Var'!R18C2:R48C2,1,0)"
ActiveCell.Copy
Range("R4:R" & vari_7).Select
ActiveSheet.Paste
Application.CutCopyMode = False




Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False

Columns("L:L").Select
    Selection.NumberFormat = "0"
    Range("L2").Select

MsgBox "METAS MONTADO-FULL", vbInformation

Sheets("Metas").Select
 Set rangodatos = Sheets("Metas").UsedRange
   rangodatos.AutoFilter Field:=21, Criteria1:="=PILOTO", Operator:=xlOr, Criteria2:="=CAV EV"
variab_9 = Sheets("Metas").Range("A" & Rows.Count).End(xlUp).Row
   Sheets("Metas").Range("A1:A" & variab_9).Copy
Sheets("CEDULAS PILOTO").Select
Range("A2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Rows(2).EntireRow.Delete

Range("D1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C[-3])-1"
variab_10 = Range("D1").Value

Range("A2").Select
vari_7 = Selection.End(xlDown).Row

Range("B2").Select


ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],HC!C3:C7,5,FALSE)"

ActiveCell.Copy
Range("B3:B" & vari_7).Select
ActiveSheet.Paste
Application.CutCopyMode = False


Sheets("Metas").Select
Range("A1").Select
ActiveSheet.ShowAllData

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Windows("PLAN_LIQ.xlsm").Activate

MsgBox "proceso FINAL FINAL completado , puede continuar con el siguiente", vbInformation
MsgBox "EL EJECUTABLE DE EXCEL SE CERRARÁ POR FAVOR ÁBRALO NUEVAMENTE Y EJECUTE EL BOTÓN DE METAS ROL-APP", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.Quit

End Sub

Sub PEGADO_1()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
On Error Resume Next
Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String




Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\FIJAS SOLO CAV")


xiaomi = "FIJAS SOLO CAV"
    ss = xiaomi & ".xlsx"

Windows(ss).Activate

Sheets("Hoja de pivote").Select

ActiveSheet.PivotTables("Tabla dinámica1").PivotSelect "", xlDataAndLabel, True
Selection.Copy


Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("A4").Select
ActiveSheet.Paste
 Application.CutCopyMode = False

Windows(ss).Activate

Sheets("Hoja1").Select
Set variable_1 = Sheets("Hoja1").Range("ak:ak").Find(What:="HPC", LookIn:=xlValues, LookAt:=xlWhole)
If Not variable_1 Is Nothing Then
Sheets("Hoja de pivote").Select
ActiveSheet.PivotTables("Tabla dinámica2").PivotSelect "", xlDataAndLabel, True
Selection.Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Fuente Hogares").Select
Range("D3").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows(ss).Activate
Sheets("Hoja1").Select
Range("A1").Select
ActiveSheet.ShowAllData
End If
If variable_1 Is Nothing Then
Windows(ss).Activate
Sheets("Hoja1").Select
Range("A1").Select
ActiveSheet.ShowAllData
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Fuente Hogares").Select
Range("D5") = "NO SE REGISTRAN VENTAS HOT"
End If
Windows(ss).Activate
Sheets("Hoja de pivote").Select
ActiveSheet.PivotTables("Tabla dinámica3").PivotSelect "", xlDataAndLabel, True
Selection.Copy


Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("G6").Select
ActiveSheet.Paste
 Application.CutCopyMode = False

Windows(ss).Activate

ActiveSheet.PivotTables("Tabla dinámica4").PivotSelect "", xlDataAndLabel, True
Selection.Copy


Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("L4").Select
ActiveSheet.Paste
 Application.CutCopyMode = False

 
 Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=False

Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "PRIMERA PARTE DEL PEGUE LISTO, CONTINÚE CON LA PARTE DOS", vbInformation
 
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation

End Sub

Sub PEGADO_2()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
On Error Resume Next

Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String




Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\FIJAS SOLO CAV")


xiaomi = "FIJAS SOLO CAV"
    ss = xiaomi & ".xlsx"

Windows(ss).Activate

Sheets("Hoja de pivote").Select

ActiveSheet.PivotTables("Tabla dinámica5").PivotSelect "", xlDataAndLabel, True
Selection.Copy



Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("P6").Select
ActiveSheet.Paste
 Application.CutCopyMode = False
 
 

 Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=False

Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "SEGUNDA PARTE DEL PEGUE LISTO, CONTINÚE CON LA PARTE TRES", vbInformation
 
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation


End Sub

Sub PEGADO_3()



Application.ScreenUpdating = False
Application.DisplayAlerts = False
On Error Resume Next

Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String




Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\FIJAS SOLO CAV")


xiaomi = "FIJAS SOLO CAV"
    ss = xiaomi & ".xlsx"

Windows(ss).Activate

Sheets("Hoja de pivote").Select

   ActiveSheet.PivotTables("Tabla dinámica6").PivotSelect "", xlDataAndLabel, True
Selection.Copy



Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("Z8").Select
ActiveSheet.Paste
 Application.CutCopyMode = False

Windows(ss).Activate

  ActiveSheet.PivotTables("Tabla dinámica7").PivotSelect "", xlDataAndLabel, True
Selection.Copy



Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("AG4").Select
ActiveSheet.Paste
 Application.CutCopyMode = False
 
 
   Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=False

Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "TERCERA PARTE DEL PEGUE LISTO, CONTINÚE CON LA PARTE CUATRO", vbInformation
 

MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA ABRIR EL ARCHIVO Y EJECUTE EL BOTÓN CUARTA PARTE", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.Quit


End Sub

Sub PEGADO_4()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
On Error Resume Next
Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Dim variable_1 As Object
Dim variable_2 As Object
Dim variable_4 As Object
Dim variable_3 As Object

Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\FIJAS SOLO CAV")


xiaomi = "FIJAS SOLO CAV"
    ss = xiaomi & ".xlsx"

Windows(ss).Activate
Sheets("Hoja de pivote").Select


 ActiveSheet.PivotTables("Tabla dinámica8").PivotSelect "", xlDataAndLabel, True
Selection.Copy



Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("AT5").Select
ActiveSheet.Paste
 Application.CutCopyMode = False



 
Windows(ss).Activate
Sheets("Hoja1").Select
Set variable_2 = Sheets("Hoja1").Range("AH:AH").Find(What:="A", LookIn:=xlValues, LookAt:=xlWhole)
If Not variable_2 Is Nothing Then
Sheets("Hoja de pivote").Select
ActiveSheet.PivotTables("Tabla dinámica9").PivotSelect "", xlDataAndLabel, True
Selection.Copy

Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("AW4").Select
ActiveSheet.Paste
Application.CutCopyMode = False
End If
If variable_2 Is Nothing Then
 Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("AW4") = "NO SE ENCONTRARON REGISTROS DE VENTAS PRE"
 End If
 
 
 
 
 Windows(ss).Activate
 Sheets("Hoja1").Select
Set variable_3 = Sheets("Hoja1").Range("aK:aK").Find(What:="PRE", LookIn:=xlValues, LookAt:=xlWhole)

If Not variable_3 Is Nothing Then
Sheets("Hoja de pivote").Select
ActiveSheet.PivotTables("Tabla dinámica10").PivotSelect "", xlDataAndLabel, True
Selection.Copy

Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("AZ3").Select
ActiveSheet.Paste
 Application.CutCopyMode = False
 End If
 If variable_3 Is Nothing Then
 
Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("AZ3") = "NO SE ENCONTRARON REGISTROS DE VENTAS PRE"
 End If
 

  Windows(ss).Activate

Sheets("Hoja1").Select
Set variable_4 = Sheets("Hoja1").Range("aK:aK").Find(What:="WINP", LookIn:=xlValues, LookAt:=xlWhole)

If Not variable_4 Is Nothing Then


'Sheets("Hoja de pivote").Select
 ' ActiveSheet.PivotTables("Tabla dinámica11").PivotSelect "", xlDataAndLabel, True
'Selection.Copy


Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
'Range("BC5").Select
'ActiveSheet.Paste
 'Application.CutCopyMode = False
 
 Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=False

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Windows("PLAN_LIQ.xlsm").Activate



MsgBox "proceso FINAL FINAL completado , puede continuar con el siguiente paso", vbInformation

 
MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA ABRIR EL ARCHIVO Y EJECUTE EL BOTON SEXTA PARTE", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.Quit
 End If
 
 If variable_4 Is Nothing Then
 Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("BC5") = "NO SE ENCONTRARON REGISTROS DE VENTAS WINP"
 
 
  Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=False

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Windows("PLAN_LIQ.xlsm").Activate



MsgBox "proceso FINAL FINAL completado , puede continuar con el siguiente paso", vbInformation
MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA ABRIR EL ARCHIVO Y EJECUTE EL BOTON SEXTA PARTE", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.Quit
 
End If
End Sub

Sub trunks()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

On Error Resume Next

Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String




Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\FIJAS SOLO CAV")


xiaomi = "FIJAS SOLO CAV"
    ss = xiaomi & ".xlsx"

Windows(ss).Activate


Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)

Worksheets.Add
ActiveSheet.Name = "Hoja de pivote"

Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("A1"))


With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("ValArpu")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields( _
        "ESTADO LEGALIZACION")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("TABLA")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica1").PivotFields("ValReg"), "Suma de ValReg", xlSum
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("ValArpu"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("ValArpu").CurrentPage _
        = "S"
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("ESTADO LEGALIZACION"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("ESTADO LEGALIZACION"). _
        CurrentPage = "LEGALIZADO"
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("TABLA").ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("TABLA").CurrentPage = _
        "PRINCIPAL"
With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("MARCA SIN TURNO")
        .Orientation = xlPageField
        .Position = 4
    End With
     ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("MARCA SIN TURNO"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("MARCA SIN TURNO")
        .PivotItems("X").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("MARCA SIN TURNO"). _
        EnableMultiplePageItems = True
        
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields( _
        "NO TIENE APP = ""X""")
        .Orientation = xlPageField
        .Position = 4
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("NO TIENE APP = ""X"""). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("NO TIENE APP = ""X"""). _
        CurrentPage = "X"
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Tipo producto")
        .Orientation = xlPageField
        .Position = 5
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Tipo producto"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Tipo producto"). _
        CurrentPage = "R"
        
        
        

'++++


Windows(ss).Activate

Sheets("Hoja de pivote").Select
Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)


Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("D1"))


With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica2").PivotFields("ValReg"), "Suma de ValReg", xlSum
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("ValArpu")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Producto Liquid")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Renta actual")
        .Orientation = xlPageField
        .Position = 3
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Renta actual"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Renta actual")
        .PivotItems("0").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Renta actual"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("ValArpu"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("ValArpu").CurrentPage _
        = "S"
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Producto Liquid"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Producto Liquid"). _
        CurrentPage = "HPC"
           With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Back_Lock")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Back_Lock"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Back_Lock"). _
        CurrentPage = "(blank)"



'++
Windows(ss).Activate

Sheets("Hoja de pivote").Select
Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)


Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("G1"))

With ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica3").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica3").PivotFields("Renta actual"), "Suma de Renta actual", xlSum
    With ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("TABLA")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Tipo venta")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("TABLA").ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("TABLA").CurrentPage = _
        "SRVADC"
    ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Tipo venta"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Tipo venta"). _
        CurrentPage = "T"
        
    With ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Producto Liquid")
        .Orientation = xlPageField
        .Position = 3
    End With
    ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Producto Liquid"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Producto Liquid")
        .PivotItems("UWARR").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Producto Liquid"). _
        EnableMultiplePageItems = True

  With ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("MARCA SIN TURNO")
        .Orientation = xlPageField
        .Position = 5
    End With
     ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("MARCA SIN TURNO"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("MARCA SIN TURNO")
        .PivotItems("X").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("MARCA SIN TURNO"). _
        EnableMultiplePageItems = True


'++

Windows(ss).Activate

Sheets("Hoja de pivote").Select
Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)


Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("J1"))

With ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica4").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica4").PivotFields("ValReg"), "Suma de ValReg", xlSum
    With ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("Producto Liquid")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("Renta actual")
        .Orientation = xlPageField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("ValArpu")
        .Orientation = xlPageField
        .Position = 3
    End With
    ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("ValArpu"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("ValArpu").CurrentPage _
        = "S"
    ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("Producto Liquid"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("Producto Liquid")
        .PivotItems("@DTH").Visible = False
        .PivotItems("100M").Visible = False
        .PivotItems("10M").Visible = False
        .PivotItems("150M").Visible = False
        .PivotItems("120M").Visible = False
        .PivotItems("80M").Visible = False
        .PivotItems("200M").Visible = False
        .PivotItems("20M").Visible = False
        .PivotItems("300M").Visible = False
        .PivotItems("30M").Visible = False
        .PivotItems("40M").Visible = False
        .PivotItems("45M").Visible = False
        .PivotItems("50M").Visible = False
        .PivotItems("5M").Visible = False
        .PivotItems("60M").Visible = False
        .PivotItems("75M").Visible = False
        .PivotItems("CV").Visible = False
        .PivotItems("DTH").Visible = False
        .PivotItems("DTHA").Visible = False
        .PivotItems("DTHS").Visible = False
        .PivotItems("HPC").Visible = False
        .PivotItems("IO").Visible = False
        .PivotItems("NVA").Visible = False
        .PivotItems("PC").Visible = False
        .PivotItems("PCI").Visible = False
        .PivotItems("PRE").Visible = False
        .PivotItems("R15").Visible = False
        .PivotItems("RJ").Visible = False
        .PivotItems("SP").Visible = False
        .PivotItems("TDB").Visible = False
        .PivotItems("TDP").Visible = False
        .PivotItems("TDS").Visible = False
        .PivotItems("TEL").Visible = False
        .PivotItems("TELDTH").Visible = False
        .PivotItems("TV").Visible = False
        .PivotItems("UW150LX2").Visible = False
        .PivotItems("UW150MX2").Visible = False
        .PivotItems("UWARR").Visible = False
        .PivotItems("VC").Visible = False
        .PivotItems("VEL").Visible = False
        .PivotItems("VEL2").Visible = False
        .PivotItems("VEL3").Visible = False
        .PivotItems("VEL4").Visible = False
        .PivotItems("VEL5").Visible = False
        .PivotItems("WIFI").Visible = False
        .PivotItems("WINP").Visible = False
        .PivotItems("HD").Visible = False
        .PivotItems("MFOX").Visible = False
        .PivotItems("15M").Visible = False
        .PivotItems("VNS").Visible = False
        .PivotItems("8M").Visible = False
        .PivotItems("180M").Visible = False
        .PivotItems("TRHO").Visible = False
        .PivotItems("120M").Visible = False
        .PivotItems("150MFO").Visible = False
        .PivotItems("180M").Visible = False
        .PivotItems("30MFO").Visible = False
        .PivotItems("45M").Visible = False
        .PivotItems("8MFO").Visible = False
        .PivotItems("45MFO").Visible = False
        .PivotItems("5MFO").Visible = False
        .PivotItems("75M").Visible = False
        .PivotItems("75MFO").Visible = False
        .PivotItems("80M").Visible = False
        .PivotItems("8M").Visible = False
        .PivotItems("TRH").Visible = False
        .PivotItems("WINPSD").Visible = False
        .PivotItems("TRHB").Visible = False
        .PivotItems("20MFO").Visible = False
        .PivotItems("60MFO").Visible = False
        .PivotItems("180MFO").Visible = False
        .PivotItems("160M").Visible = False
        .PivotItems("240M").Visible = False
        .PivotItems("400M").Visible = False
        .PivotItems("140M").Visible = False
        .PivotItems("200MFO").Visible = False
        .PivotItems("210M").Visible = False
        .PivotItems("23M").Visible = False
        .PivotItems("300MFO").Visible = False
        .PivotItems("310M").Visible = False
        .PivotItems("40MFO").Visible = False
        .PivotItems("50MFO").Visible = False
        .PivotItems("160M").Visible = False
        .PivotItems("240M").Visible = False
        .PivotItems("400M").Visible = False
        .PivotItems("HD").Visible = False
        .PivotItems("DTHPRE").Visible = False
        .PivotItems("(blank)").Visible = False
        .PivotItems("100MFO").Visible = False
        .PivotItems("UW75L").Visible = False
        .PivotItems("WINSD").Visible = False
              
    End With
    ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("Producto Liquid"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("Renta actual"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("Renta actual")
        .PivotItems("0").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("Renta actual"). _
        EnableMultiplePageItems = True

   With ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("Back_Lock")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("Back_Lock"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica4").PivotFields("Back_Lock"). _
        CurrentPage = "(blank)"


'++
Windows(ss).Activate

Sheets("Hoja de pivote").Select
Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)


Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("M1"))

  With ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica5").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica5").PivotFields("ValReg"), "Suma de ValReg", xlSum
    With ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("TABLA")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("MARCA SIN TURNO")
        .Orientation = xlPageField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("Tabla dinámica5").PivotFields( _
        "ESTADO LEGALIZACION")
        .Orientation = xlPageField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("ValArpu")
        .Orientation = xlPageField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("Renta actual")
        .Orientation = xlPageField
        .Position = 5
    End With
    With ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("Producto Liquid")
        .Orientation = xlPageField
        .Position = 6
    End With
    ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("Producto Liquid"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("Producto Liquid"). _
        CurrentPage = "VEL2"
    ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("Renta actual"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("Renta actual")
        .PivotItems("0").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("Renta actual"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("ValArpu"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("ValArpu").CurrentPage _
        = "S"
    Range("B4").Select
    ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("ESTADO LEGALIZACION"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("ESTADO LEGALIZACION"). _
        CurrentPage = "LEGALIZADO"
    ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("MARCA SIN TURNO"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("MARCA SIN TURNO")
        .PivotItems("X").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("MARCA SIN TURNO"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("TABLA").ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica5").PivotFields("TABLA").CurrentPage = _
        "PRINCIPAL"




'al poco que debutó marado marado
'++
Windows(ss).Activate

Sheets("Hoja de pivote").Select
Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)

Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("V1"))


 With ActiveSheet.PivotTables("Tabla dinámica6").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica6").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica6").PivotFields("Cant_multi"), "Cuenta de Cant_multi", xlCount
    With ActiveSheet.PivotTables("Tabla dinámica6").PivotFields( _
        "Cuenta de Cant_multi")
        .Caption = "Suma de Cant_multi"
        .Function = xlSum
    End With
    With ActiveSheet.PivotTables("Tabla dinámica6").PivotFields("TABLA")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica6").PivotFields("Back_Lock")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica6").PivotFields("TABLA").CurrentPage = _
        "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica6").PivotFields("TABLA")
        .PivotItems("MIGRA").Visible = False
        .PivotItems("MOVIL").Visible = False
        .PivotItems("PRINCIPAL").Visible = False
        .PivotItems("SRVADC").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica6").PivotFields("TABLA"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("Tabla dinámica6").PivotFields("Back_Lock"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica6").PivotFields("Back_Lock"). _
        CurrentPage = "(blank)"



'++
Windows(ss).Activate

Sheets("Hoja de pivote").Select
Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)


Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("AB1"))



    ActiveSheet.PivotTables("Tabla dinámica7").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica7").PivotFields("Conver_Lider"), "Suma de Conver_Lider", xlSum
    With ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("Conver_Lider")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("Back_Lock")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("No pago Est")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("TABLA")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("ValArpu")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("Renta actual")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("Conver_Lider"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("Conver_Lider"). _
        CurrentPage = "1"
    ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("Back_Lock"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("Back_Lock"). _
        CurrentPage = "(blank)"
    ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("No pago Est"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("No pago Est")
        .PivotItems("0").Visible = False
        .PivotItems("1").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("No pago Est"). _
        EnableMultiplePageItems = True
    Range("B4").Select
    ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("TABLA").ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("TABLA").CurrentPage = _
        "PRINCIPAL"
    ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("ValArpu"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("ValArpu").CurrentPage _
        = "S"
    ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("Renta actual"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("Renta actual")
        .PivotItems("0").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica7").PivotFields("Renta actual"). _
        EnableMultiplePageItems = True
    
        
        

'++
Windows(ss).Activate

Sheets("Hoja de pivote").Select
Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)


Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("AE1"))

    With ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica8").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica8").PivotFields("ValReg"), "Suma de ValReg", xlSum
    With ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("ValArpu")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("Tipo venta")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("Producto Liquid")
        .Orientation = xlPageField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("Renta actual")
        .Orientation = xlPageField
        .Position = 3
    End With
    ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("ValArpu"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("ValArpu").CurrentPage _
        = "S"
    ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("Renta actual"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("Renta actual")
        .PivotItems("0").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("Renta actual"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("Producto Liquid"). _
        CurrentPage = "(All)"
   
   ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("Producto Liquid"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("Producto Liquid"). _
        CurrentPage = "UWARR"
    ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("Producto Liquid"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("Tipo venta"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("Tipo venta"). _
        CurrentPage = "T"
    With ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("Back_Lock")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("Back_Lock"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica8").PivotFields("Back_Lock"). _
        CurrentPage = "(blank)"

'++
Windows(ss).Activate

Sheets("Hoja de pivote").Select
Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)


Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("AH1"))
 
 
    With ActiveSheet.PivotTables("Tabla dinámica9").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica9").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica9").PivotFields("ValReg"), "Suma de ValReg", xlSum
    With ActiveSheet.PivotTables("Tabla dinámica9").PivotFields("Tipo venta")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica9").PivotFields("ValArpu")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica9").PivotFields("Renta actual")
        .Orientation = xlPageField
        .Position = 2
    End With
    ActiveSheet.PivotTables("Tabla dinámica9").PivotFields("Tipo venta"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica9").PivotFields("Tipo venta"). _
        CurrentPage = "A"
    ActiveSheet.PivotTables("Tabla dinámica9").PivotFields("ValArpu"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica9").PivotFields("ValArpu").CurrentPage _
        = "S"
    ActiveSheet.PivotTables("Tabla dinámica9").PivotFields("Renta actual"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica9").PivotFields("Renta actual")
        .PivotItems("0").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica9").PivotFields("Renta actual"). _
        EnableMultiplePageItems = True

   With ActiveSheet.PivotTables("Tabla dinámica9").PivotFields("Back_Lock")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica9").PivotFields("Back_Lock"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica9").PivotFields("Back_Lock"). _
        CurrentPage = "(blank)"

 '++
Windows(ss).Activate

Sheets("Hoja de pivote").Select
Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)


Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("AK1"))
 
 
 

   With ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica10").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica10").PivotFields("ValReg"), "Suma de ValReg", xlSum
    With ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("Producto Liquid")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("Renta actual")
        .Orientation = xlPageField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("ValArpu")
        .Orientation = xlPageField
        .Position = 3
    End With
    ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("ValArpu"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("ValArpu").CurrentPage _
        = "S"
    ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("Producto Liquid"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("Producto Liquid"). _
        CurrentPage = "PRE"
    ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("Renta actual"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("Renta actual")
        .PivotItems("0").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("Renta actual"). _
        EnableMultiplePageItems = True

   With ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("Back_Lock")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("Back_Lock"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("Back_Lock"). _
        CurrentPage = "(blank)"

'
'
'

'++
Windows(ss).Activate

Sheets("Hoja de pivote").Select
Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)


Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("AN1"))



       With ActiveSheet.PivotTables("Tabla dinámica11").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica11").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica11").PivotFields("ValReg"), "Suma de ValReg", xlSum
    With ActiveSheet.PivotTables("Tabla dinámica11").PivotFields("Producto Liquid")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica11").PivotFields("Renta actual")
        .Orientation = xlPageField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("Tabla dinámica11").PivotFields("ValArpu")
        .Orientation = xlPageField
        .Position = 3
    End With
    Range("B1").Select
    ActiveSheet.PivotTables("Tabla dinámica11").PivotFields("ValArpu"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica11").PivotFields("ValArpu").CurrentPage _
        = "S"
    ActiveSheet.PivotTables("Tabla dinámica11").PivotFields("Renta actual"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica11").PivotFields("Renta actual")
        .PivotItems("0").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica11").PivotFields("Renta actual"). _
        EnableMultiplePageItems = True
  
  ActiveSheet.PivotTables("Tabla dinámica11").PivotFields("Producto Liquid"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica11").PivotFields("Producto Liquid")
        .PivotItems("@DTH").Visible = False
        .PivotItems("100M").Visible = False
        .PivotItems("10M").Visible = False
        .PivotItems("120M").Visible = False
        .PivotItems("150M").Visible = False
        .PivotItems("150MFO").Visible = False
        .PivotItems("15M").Visible = False
        .PivotItems("180M").Visible = False
        .PivotItems("200M").Visible = False
        .PivotItems("300M").Visible = False
        .PivotItems("30M").Visible = False
        .PivotItems("30MFO").Visible = False
        .PivotItems("40M").Visible = False
        .PivotItems("45M").Visible = False
        .PivotItems("45MFO").Visible = False
        .PivotItems("50M").Visible = False
        .PivotItems("5M").Visible = False
        .PivotItems("5MFO").Visible = False
        .PivotItems("60M").Visible = False
        .PivotItems("75M").Visible = False
        .PivotItems("75MFO").Visible = False
        .PivotItems("80M").Visible = False
        .PivotItems("8M").Visible = False
        .PivotItems("8MFO").Visible = False
        .PivotItems("CV").Visible = False
        .PivotItems("DTH").Visible = False
        .PivotItems("DTHA").Visible = False
        .PivotItems("DTHS").Visible = False
        .PivotItems("FOX").Visible = False
        .PivotItems("HBO").Visible = False
        .PivotItems("HPC").Visible = False
        .PivotItems("IO").Visible = False
        .PivotItems("NVA").Visible = False
        .PivotItems("PC").Visible = False
        .PivotItems("PCI").Visible = False
        .PivotItems("R15").Visible = False
        .PivotItems("RJ").Visible = False
        .PivotItems("SP").Visible = False
        .PivotItems("TDB").Visible = False
        .PivotItems("TDP").Visible = False
        .PivotItems("TDS").Visible = False
        .PivotItems("TEL").Visible = False
        .PivotItems("TELDTH").Visible = False
        .PivotItems("TRHB").Visible = False
        .PivotItems("TRHO").Visible = False
        .PivotItems("TV").Visible = False
        .PivotItems("UWARR").Visible = False
        .PivotItems("VC").Visible = False
        .PivotItems("VEL").Visible = False
        .PivotItems("VEL2").Visible = False
        .PivotItems("VEL3").Visible = False
        .PivotItems("20M").Visible = False
        .PivotItems("20MFO").Visible = False
        .PivotItems("PRE").Visible = False
        .PivotItems("GOLD").Visible = False
        .PivotItems("HD").Visible = False
        .PivotItems("60MFO").Visible = False
        .PivotItems("180MFO").Visible = False
        .PivotItems("160M").Visible = False
        .PivotItems("240M").Visible = False
        .PivotItems("400M").Visible = False
        .PivotItems("HD").Visible = False
        .PivotItems("MFOX").Visible = False
        .PivotItems("15M").Visible = False
        .PivotItems("VNS").Visible = False
        .PivotItems("8M").Visible = False
        .PivotItems("180M").Visible = False
        .PivotItems("TRHO").Visible = False
        .PivotItems("120M").Visible = False
        .PivotItems("150MFO").Visible = False
        .PivotItems("180M").Visible = False
        .PivotItems("30MFO").Visible = False
        .PivotItems("45M").Visible = False
        .PivotItems("8MFO").Visible = False
        .PivotItems("45MFO").Visible = False
        .PivotItems("5MFO").Visible = False
        .PivotItems("75M").Visible = False
        .PivotItems("75MFO").Visible = False
        .PivotItems("80M").Visible = False
        .PivotItems("8M").Visible = False
        .PivotItems("TRH").Visible = False
        .PivotItems("TRHB").Visible = False
        .PivotItems("20MFO").Visible = False
        .PivotItems("60MFO").Visible = False
        .PivotItems("180MFO").Visible = False
        .PivotItems("160M").Visible = False
        .PivotItems("240M").Visible = False
        .PivotItems("400M").Visible = False
        .PivotItems("140M").Visible = False
        .PivotItems("200MFO").Visible = False
        .PivotItems("210M").Visible = False
        .PivotItems("23M").Visible = False
        .PivotItems("300MFO").Visible = False
        .PivotItems("310M").Visible = False
        .PivotItems("40MFO").Visible = False
        .PivotItems("50MFO").Visible = False
        .PivotItems("160M").Visible = False
        .PivotItems("240M").Visible = False
        .PivotItems("400M").Visible = False
        .PivotItems("HD").Visible = False
        .PivotItems("DTHPRE").Visible = False
        .PivotItems("(blank)").Visible = False
        
        
        
        
    End With
    With ActiveSheet.PivotTables("Tabla dinámica11").PivotFields("Producto Liquid")
        .PivotItems("VEL4").Visible = False
        .PivotItems("VEL5").Visible = False
        .PivotItems("WIFI").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica11").PivotFields("Producto Liquid"). _
        EnableMultiplePageItems = True

    With ActiveSheet.PivotTables("Tabla dinámica11").PivotFields("Back_Lock")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica11").PivotFields("Back_Lock"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica11").PivotFields("Back_Lock"). _
        CurrentPage = "(blank)"
 
 
 Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=True

Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "proceso FINAL FINAL completado ,puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
End Sub


Sub pervariable_5()
Application.ScreenUpdating = False
Application.DisplayAlerts = False



Dim archivos

Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\HC\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\HC\" & archivos

archivos = Dir
Loop
'
vari_1 = ActiveWorkbook.Name
'
tt = vari_1





Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\PLANTILLA_HC")



xiaomi = "PLANTILLA_HC"
    ss = xiaomi & ".xlsx"


Windows(tt).Activate

 variable_5 = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row

 Sheets(1).Range("A1:U" & variable_5).Copy


    Windows(ss).Activate

Sheets("HC").Select
Range("A2").Select
ActiveSheet.Paste



Windows(tt).Activate
Application.DisplayAlerts = False
ActiveWorkbook.Close
Application.DisplayAlerts = True

MsgBox "proceso de pegue completado", vbInformation


Windows(ss).Activate
Sheets("HC").Select
    Set rangodatos = Sheets("HC").UsedRange
    rangodatos.AutoFilter Field:=19, Criteria1:="<>"
    
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=RC[14]"
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
     Range("E1") = "Cod Cargo"


Range("G1").Select
    ActiveCell.FormulaR1C1 = "=RC[13]"
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
     Range("G1") = "Sueldo Variable"
     ActiveSheet.ShowAllData

ActiveWorkbook.Save
ActiveWorkbook.Close

Application.ScreenUpdating = True
Application.DisplayAlerts = True


MsgBox "proceso FINAL FINAL completado , puede continuar con el siguiente paso", vbInformation


End Sub

Sub variable_6()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim archivos

Dim vari_2 As Excel.Workbook
Dim vari_3 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\AUSENTISMOS\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\AUSENTISMOS\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate
Sheets(1).Select
    Set rangodatos = Sheets(1).UsedRange
    rangodatos.AutoFilter Field:=5, Criteria1:="=ret*", _
        Operator:=xlAnd

Range("E1").Select
 With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell = "retiro"
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
     Range("E1") = "CLASE AUSENTISMO"
     ActiveSheet.ShowAllData
   
   MsgBox "listos retiros", vbInformation
   
Range("I1") = "HC ACTUAL"
   


Windows(tt).Activate


Range("M1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C[-12])-1"


Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\PLANTILLA_HC")


'
xiaomi = "PLANTILLA_HC"
    ss = xiaomi & ".xlsx"

Windows(ss).Activate

Windows(tt).Activate

 va_1 = Range("M1").Value
 Range("I2").Select
 For i = 1 To va_1

 ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-8],[PLANTILLA_HC.xlsx]HC!C1:C5,5,FALSE)"
    
   ActiveCell.Offset(1, 0).Select
    Next i


Range("J1") = "HC VIEJO"
Windows(ss).Activate
ActiveWorkbook.Save
ActiveWorkbook.Close

MsgBox "LISTOS EL ACTUAL, ", vbInformation




Set vari_3 = Workbooks.Open("D:\AUTOMATIZACION\VIEJA\HC_VIEJO")


'
variable_ = "HC_VIEJO"
    variab_10 = variable_ & ".xlsx"

Windows(variab_10).Activate


Windows(tt).Activate
Sheets(1).Select
carita = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Range("A1").Select
Selection.AutoFilter
 ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:J" & carita), , xlYes).Name = _
        "Tabla1"
    Range("Tabla1[#All]").Select

   

Range("j2").Select
 For i = 1 To va_1

ActiveCell.FormulaR1C1 = "=VLOOKUP([@CEDULA],[HC_VIEJO.xlsx]HC!C3:C7,5,FALSE)"
    
   ActiveCell.Offset(1, 0).Select
    Next i

MsgBox "listo viejo", vbInformation



Range("Tabla1").AutoFilter Field:=9, Criteria1:="#N/A"
Range("Tabla1").AutoFilter Field:=10, Criteria1:="<>#N/A" _
        , Operator:=xlAnd


Range("A1").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy


Windows(variab_10).Activate
Worksheets.Add.Name = "NUEVA"
Sheets("NUEVA").Select
Range("A1").Select
ActiveSheet.Paste

Windows(tt).Activate

ActiveWorkbook.Save
ActiveWorkbook.Close

Windows(variab_10).Activate
Sheets("HC").Select

Range("AI1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C[-32])-1"

 
CISCO = Range("AI1").Value

Range("AB2").Select
 For i = 1 To CISCO

ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-25],NUEVA!C[-27],1,FALSE)"
    
   ActiveCell.Offset(1, 0).Select
    Next i



Sheets("HC").Select
    Set rangodatos = Sheets("HC").UsedRange
    rangodatos.AutoFilter Field:=28, Criteria1:="<>#N/A" _
    , Operator:=xlAnd

 rangodatos.AutoFilter Field:=7, Criteria1:=Array("Consultor Servicio Personalizado A Clientes", "Consultor Integral Servicio A Clientes", "Asesor Servicio Al Cliente", "Consultor Integral Servicio A Clientes Sr", "Asesor Integral Servicio Al Cliente", "Consultor(a) Integral Servicio A Clientes", "Asesor(a) Servicio Al Cliente", "Consultor(a) Integral Servicio A Clientes Sr", "Consultor(a) Servicio Personalizado A Clientes", "Asesor(a) Integral Servicio Al Cliente"), _
        Operator:=xlFilterValues

rangodatos.AutoFilter Field:=9, Criteria1:="<>0" _
    , Operator:=xlAnd

Range("C1").Select

Range(Selection, Selection.End(xlDown)).Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
dedo = Sheets("HC").Range("C" & Rows.Count).End(xlUp).Row
Sheets("HC").Select
Sheets("HC").Range("C1:W" & dedo).Copy

Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\PLANTILLA_HC")


'
xiaomi = "PLANTILLA_HC"
    ss = xiaomi & ".xlsx"

Windows(ss).Activate

enero = Sheets("HC").Range("A" & Rows.Count).End(xlUp).Row

Sheets("HC").Select
Range("A" & enero + 1).Select
ActiveSheet.Paste


Sheets("HC").Select
    Set rangodatos = Sheets("HC").UsedRange
    rangodatos.AutoFilter Field:=1, Criteria1:="Expediente"
    
With Range("A1")

Range(Cells(Rows.Count, .Column).End(xlUp), Cells(.Row + 1, Columns.Count).End(xlToLeft)).SpecialCells(12).Delete

End With

Range("A1").Select
ActiveSheet.ShowAllData

Windows(ss).Activate
ActiveWorkbook.Save
ActiveWorkbook.Close
Windows(variab_10).Activate
ActiveWorkbook.Save
ActiveWorkbook.Close

Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "proceso FINAL FINAL completado , puede continuar con el siguiente paso", vbInformation



End Sub


Sub UMBRAL()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos

Dim vari_2 As Excel.Workbook

Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\UMBRAL\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\UMBRAL\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate
Sheets("Data").Select

Range("O1") = "CRUCE POR USUARIO"


archi = Dir("D:\AUTOMATIZACION\NPS\USUARIO\*.csv")

Do While archi <> ""
Workbooks.Open "D:\AUTOMATIZACION\NPS\USUARIO\" & archi

archi = Dir
Loop

vari_3 = ActiveWorkbook.Name

dell = vari_3

Windows(dell).Activate

Set rangodatos = Sheets(1).UsedRange
    rangodatos.AutoFilter Field:=3, Criteria1:=0
Range("D1") = "NO REGISTRA"
Range("D1").Select
 Selection.Copy
    ActiveCell.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    ActiveSheet.Paste
    Range("D1").Select
    Application.CutCopyMode = False
Range("D1") = "VAL_NPS_CALCULADO"
Range("A1").Select
ActiveSheet.ShowAllData
    MsgBox "LISTO ENCUESTAS", vbInformation


Windows(tt).Activate

Range("p1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C[-15])-1"
vari_5 = Range("p1").Value
Range("o2").Select

For i = 1 To vari_5

ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-3],NPS_USUARIORED.csv!C2:C4,3,FALSE),""NO REGISTRA"")"

ActiveCell.Offset(1, 0).Select

Next i

 Columns("O:O").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("O1").Select

MsgBox "listo umbral por usuario", vbInformation


Windows(dell).Activate
ActiveWorkbook.Close SaveChanges:=False


Arc = Dir("D:\AUTOMATIZACION\NPS\*.csv")

Do While Arc <> ""
Workbooks.Open "D:\AUTOMATIZACION\NPS\" & Arc

Arc = Dir
Loop

va_2 = ActiveWorkbook.Name

va_3 = va_2

Windows(va_3).Activate


Set rangodatos = Sheets(1).UsedRange
    rangodatos.AutoFilter Field:=3, Criteria1:=0
Range("D1") = "NO REGISTRA"
Range("D1").Select
 Selection.Copy
    ActiveCell.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    ActiveSheet.Paste
    Range("D1").Select
    Application.CutCopyMode = False
Range("D1") = "VAL_NPS_CALCULADO"
Range("A1").Select
ActiveSheet.ShowAllData
    MsgBox "LISTO ENCUESTAS", vbInformation




Windows(tt).Activate
Range("Q1") = "CRUCE POR CEDULA"
Range("Q2").Select
For i = 1 To vari_5

ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-16],NPS_CEDULA.csv!C2:C4,3,FALSE),""NO REGISTRA"")"

ActiveCell.Offset(1, 0).Select

Next i

 Columns("Q:Q").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("Q1").Select
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=True
MsgBox "listo umbral por CEDULA", vbInformation


Windows(va_3).Activate
Sheets(1).Select
    Set rangodatos = Sheets(1).UsedRange
    rangodatos.AutoFilter Field:=4, Criteria1:="<>0", _
        Operator:=xlAnd, Criteria2:="<>NO REGISTRA"
    
va_4 = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row

   Sheets(1).Range("B1:D" & va_4).Copy


Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\NPS\NPS_CED_PLANTILLA")
xiaomi = "NPS_CED_PLANTILLA"
    ss = xiaomi & ".xlsx"

Windows(ss).Activate

Sheets(1).Select
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows(va_3).Activate
ActiveWorkbook.Close SaveChanges:=False
Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=True
MsgBox "LISTOS faltantes umbral", vbInformation
MsgBox "proceso FINAL FINAL completado , puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub



Sub vacaciones()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos

Dim vari_2 As Excel.Workbook
Dim vari_3 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\VACACIONES\*.csv")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\VACACIONES\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate

Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1)), TrailingMinusNumbers:=True
    Range("A1").Select



Sheets(1).Select
Range("A1").Select

col = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("A1:G" & col).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Ausentismos-Vaca-Umb").Select
Range("K2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False

MsgBox "LISTO VACAS, ", vbInformation


Range("AN1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C[-39])-1"
vari_5 = Range("AN1").Value
Range("I2").Select

For i = 1 To vari_5

  ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-8],HC!C3:C7,5,FALSE)"

ActiveCell.Offset(1, 0).Select

Next i


Range("AO1").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(C[-10])-1"
dell = Range("AO1").Value
Range("AK2").Select



For i = 1 To dell

ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-6],HC!C3:C9,7,FALSE)"

ActiveCell.Offset(1, 0).Select

Next i

MsgBox "listos validadores", vbInformation

archivos = Dir("D:\AUTOMATIZACION\GARANTIZADOS\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\GARANTIZADOS\" & archivos

archivos = Dir
Loop

vari_4 = ActiveWorkbook.Name

pp = vari_4

Windows(pp).Activate


Sheets(1).Select
Range("A1").Select

las = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("A2:E" & las).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Garantizado").Select
Range("A2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

Windows(pp).Activate
ActiveWorkbook.Close SaveChanges:=False



Range("H1").Select
     ActiveCell.FormulaR1C1 = "=COUNTA(C[-7])-1"
claro = Range("H1").Value
Range("F2").Select



For i = 1 To claro
 ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],HC!C3:C7,5,FALSE)"

ActiveCell.Offset(1, 0).Select

Next i


MsgBox "proceso FINAL FINAL completado , puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True


End Sub


Sub vari_10()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Sheets("Meses").Select
variable_5 = Range("M1").Value
variab_10 = Range("M2").Value




Windows("PLAN_LIQ.xlsm").Activate

Sheets("HC").Visible = False
Sheets("CEDULAS PILOTO").Protect Password:="TOP"
Sheets("Metas").Protect Password:="TOP"
Sheets("INCENTIVO").Protect Password:="TOP"
'Sheets("INCENTIVO MOTOROLA - CLARO").Protect Password:="TOP"
'Sheets("INCENTIVO POSPAGO").Protect Password:="TOP"

Sheets("AISLAMIENTO COVID").Select
Columns("C:C").Select
    Selection.EntireColumn.Hidden = True

Sheets("AISLAMIENTO COVID").Protect Password:="TOP"

Sheets("Resumen_plan_power").Protect Password:="TOP"





Sheets("Liquidacion").Select
Columns("D:D").Select
    Selection.EntireColumn.Hidden = True

Sheets("Liquidacion").Protect Password:="TOP"


Sheets("Consultores_Inspira").Select
Columns("D:D").Select
    Selection.EntireColumn.Hidden = True

Sheets("Consultores_Inspira").Protect Password:="TOP"


Sheets("Liquidacion_PART_TIME").Select
va_6 = Range("A1").Value

If va_6 = 0 Then
Sheets("Liquidacion_PART_TIME").Visible = False
Else
Sheets("Liquidacion_PART_TIME").Select
Columns("D:D").Select
    Selection.EntireColumn.Hidden = True

Sheets("Liquidacion_PART_TIME").Protect Password:="TOP"

End If

Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Select
Columns("D:D").Select
    Selection.EntireColumn.Hidden = True

Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Protect Password:="TOP"
Sheets("Resumen Puntos").Protect Password:="TOP"
Sheets("MOVIL").Protect Password:="TOP"

Sheets("Fuente Hogares").Protect Password:="TOP"
Sheets("Tabla Var").Protect Password:="TOP"
Sheets("Meses").Protect Password:="TOP"
Sheets("ROLES_EXPERIENCIA AL CLIENTE").Protect Password:="TOP"
Sheets("NPS-UMBRAL").Select
Columns("B:B").Select
    Selection.EntireColumn.Hidden = True

Sheets("NPS-UMBRAL").Protect Password:="TOP"

Sheets("Ausentismos-Vaca-Umb").Select
Columns("C:C").Select
    Selection.EntireColumn.Hidden = True
Columns("L:L").Select
    Selection.EntireColumn.Hidden = True
Columns("AF:AF").Select
    Selection.EntireColumn.Hidden = True
    
    
Sheets("Ausentismos-Vaca-Umb").Protect Password:="TOP"
Sheets("Garantizado").Protect Password:="TOP"


Sheets("VENTAS_OTRAS GESTIONES").Protect Password:="TOP"

Sheets("OTRAS GESTIONES").Select
Columns("D:D").Select
    Selection.EntireColumn.Hidden = True
Sheets("OTRAS GESTIONES").Protect Password:="TOP"

Sheets("ASESORES TMK").Select
sancio = Range("A1").Value

If sancio = 0 Then
Sheets("ASESORES TMK").Visible = False
Else
Sheets("ASESORES TMK").Select
Columns("D:D").Select
    Selection.EntireColumn.Hidden = True

Sheets("ASESORES TMK").Protect Password:="TOP"

End If


Sheets("DESARROLLO+PROYECTOS").Select
Columns("D:D").Select
    Selection.EntireColumn.Hidden = True
Sheets("DESARROLLO+PROYECTOS").Protect Password:="TOP"



Sheets("CUMPLIMIENTOS").Select
Columns("B:B").Select
    Selection.EntireColumn.Hidden = True
Sheets("CUMPLIMIENTOS").Protect Password:="TOP"


Sheets("Desarrollo").Select
Columns("B:B").Select
    Selection.EntireColumn.Hidden = True
Sheets("Desarrollo").Protect Password:="TOP"



Sheets("Proyectos").Select
Columns("A:A").Select
    Selection.EntireColumn.Hidden = True
Sheets("Proyectos").Protect Password:="TOP"



Sheets("ROL-APP").Select
claro = Range("A1").Value
If claro = 0 Then
Sheets("ROL-APP").Visible = False
Else
Sheets("ROL-APP").Select
Columns("D:D").Select
    Selection.EntireColumn.Hidden = True
Sheets("ROL-APP").Protect Password:="TOP"
End If
Sheets("METAS TMK").Protect Password:="TOP"
'Sheets("PROMEDIOS_CONVERGENCIA").Protect Password:="TOP"
'Sheets("PROMEDIOS_PROTECCION_INGRESO").Protect Password:="TOP"

'Sheets("% CUM_4 MES").Select
'Columns("B:B").Select
'    Selection.EntireColumn.Hidden = True
'Sheets("% CUM_4 MES").Protect Password:="TOP"



ActiveWorkbook.Protect ("TOP")

  RUTA = "D:\AUTOMATIZACION\LIQ_FINAL\"
    GENERAR_ARCHIVO_EXCEL = RUTA & "Liq_" & variable_5 & variab_10 & "_Consultor Integral Servicio al Cliente_VAL" & ".xlsx"

   ActiveWorkbook.SaveAs Filename:= _
        GENERAR_ARCHIVO_EXCEL, FileFormat:= _
        xlOpenXMLWorkbook, Password:="1670", CreateBackup:=False


simpson = ActiveWorkbook.Name

hh = simpson

Windows(hh).Activate
ActiveWorkbook.Close SaveChanges:=True


Application.Calculation = xlCalculationAutomatic



MsgBox "FELICIDADES ...UNA LIQUIDACIÓN QUE SE DEMORABA DOS DÍAS LA ACABASTE DE HACER EN 1 HORA", vbInformation
Application.Calculation = xlCalculationAutomatic
ActiveWorkbook.Close SaveChanges:=True
Application.Quit
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub


Sub atajitos()
UserForm1.Show
End Sub

Sub CLARITO()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos

Dim vari_2 As Excel.Workbook

Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\METAS\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\METAS\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate

Sheets("Fuente_Consultor App").Select
vegas = Sheets("Fuente_Consultor App").Range("A" & Rows.Count).End(xlUp).Row
Sheets("Fuente_Consultor App").Range("B2:L" & vegas).Copy
    
    
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("Y3").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("Y3").Select
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
Application.ScreenUpdating = True
Application.DisplayAlerts = True
 
 MsgBox "Listo primera fase", vbInformation
End Sub
Sub quinteto()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

On Error Resume Next

Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Dim variable_1 As Object




Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\FIJAS SOLO CAV")


xiaomi = "FIJAS SOLO CAV"
    ss = xiaomi & ".xlsx"




Windows(ss).Activate

Sheets("Hoja de pivote").Select
Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)
Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("AS1"))


 With ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica12").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica12").PivotFields("ValReg"), "Suma de ValReg", xlSum
    With ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("Producto Liquid")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("Renta actual")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("ValArpu")
        .Orientation = xlPageField
        .Position = 2
    End With
    ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("Producto Liquid"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("Producto Liquid"). _
        CurrentPage = "MFOX"
    ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("ValArpu"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("ValArpu").CurrentPage _
        = "S"
    ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("Renta actual"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("Renta actual")
        .PivotItems("0").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("Renta actual"). _
        EnableMultiplePageItems = True

'++

Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)

Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("AV1"))


 With ActiveSheet.PivotTables("Tabla dinámica13").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica13").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica13").PivotFields("ValReg"), "Suma de ValReg", xlSum
    With ActiveSheet.PivotTables("Tabla dinámica13").PivotFields("Producto Liquid")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica13").PivotFields("Renta actual")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica13").PivotFields("ValArpu")
        .Orientation = xlPageField
        .Position = 2
    End With
    ActiveSheet.PivotTables("Tabla dinámica13").PivotFields("Producto Liquid"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica13").PivotFields("Producto Liquid"). _
        CurrentPage = "MHBO"
    ActiveSheet.PivotTables("Tabla dinámica13").PivotFields("ValArpu"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica13").PivotFields("ValArpu").CurrentPage _
        = "S"
    ActiveSheet.PivotTables("Tabla dinámica13").PivotFields("Renta actual"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica13").PivotFields("Renta actual")
        .PivotItems("0").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica13").PivotFields("Renta actual"). _
        EnableMultiplePageItems = True


Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=True
MsgBox "listo tablas dinámicas", vbInformation
Call CAPRI
Call SAGAN

MsgBox "LISTO EL CUARTO PEGUE, CONTINÚE CON LA PARTE CINCO", vbInformation


Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub shimano()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos

Dim vari_2 As Excel.Workbook

Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\TRANSACCIONES_APP\*.xlsx")
If archivos = "" Then
MsgBox "NO HAY NINGÚN ARCHIVO DE TRANSACCIONES APP CARGADO", vbCritical
Else
Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\TRANSACCIONES_APP\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate

Sheets("Turnos_Consultores_App_y_TP").Select
vegas = Sheets("Turnos_Consultores_App_y_TP").Range("A" & Rows.Count).End(xlUp).Row
Sheets("Turnos_Consultores_App_y_TP").Range("A2:B" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("NPS-UMBRAL").Select
Range("W2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A2").Select
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End If
MsgBox "Proceso FINAL FINAL Completado, puede continuar con el siguiente paso", vbInformation
End Sub
 Sub VALI_APP()
 Application.ScreenUpdating = False
Application.DisplayAlerts = False
 Dim va_5 As String
 
 Sheets("HC").Select
 va_5 = Range("AC1").Value
 Range("z2").Select
 For i = 1 To va_5
 ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-23],Metas!C25,1,FALSE)"
 ActiveCell.Offset(1, 0).Select
 Next i
 MsgBox "listos HC, con APP", vbInformation
 Application.ScreenUpdating = True
Application.DisplayAlerts = True
 End Sub
Sub inventos()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("ROL-APP").Select
Range("I2") = UserForm7.ComboBox1
Range("a1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C[2])-2"
vari_9 = Range("a1").Value
Range("c4").Select

For i = 1 To vari_9

ActiveCell.Offset(0, 1).Range("A1").Select

On Error GoTo marcador
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],HC!C3:C5,3,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],HC!C3:C7,5,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-3],Metas!C25:C27,3,0),""Sin Oficina"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-4],Metas!C25:C28,4,0),""Sin Oficina"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(VLOOKUP(RC[-5],HC!C3:C9,7,0),0)>0,VLOOKUP(RC[-5],'DESARROLLO+PROYECTOS'!C3:C9,7,FALSE),IFERROR(VLOOKUP(RC[-5],HC!C3:C9,7,0),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C9,Meses!C[-8]:C[-7],2,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((IF((IF(AND(RC[40]=""retiro"",RC[37]=RC[35],RC[38]=RC[-1]),IF((IF(RC[36]>(CONCATENATE(RC[35],RC[-1],RC[34])-RC[32]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[34]-RC[39])-RC[32])+1,(RC[34]-RC[39]+1)+(RC[34]-(RC[32]+RC[33]))),IF(AND(RC[37]=RC[35],RC[38]=RC[-1]),(RC[34]-RC[39]+1)-(RC[32]+RC[33]),RC[34]-(RC[32]+RC[33]))))<0,0,IF(AND(RC[40]=""retiro"",RC[37]=RC[35],RC[38]=RC[-1]),IF((IF(RC[36]>(CONCATENATE(RC[35],RC[-1],RC[34])-RC[32]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[34]-RC[39])-RC[32])+1,(RC[34]-RC[39]+1)+(RC[34]-(RC[32]+RC[33]))),IF(AND(RC[37]=RC[35],RC[38]=RC[-1]),(RC[34]-RC[39]+1)-(RC[32]+RC[33]),RC[34]-(RC[32]+RC[33])))-IFERROR(VLOOKUP(RC[-7],'AISLAMIENTO COVID'!C1:C7,7,FALSE),0)))<0,0," & _
        "IF((IF(AND(RC[40]=""retiro"",RC[37]=RC[35],RC[38]=RC[-1]),IF((IF(RC[36]>(CONCATENATE(RC[35],RC[-1],RC[34])-RC[32]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[34]-RC[39])-RC[32])+1,(RC[34]-RC[39]+1)+(RC[34]-(RC[32]+RC[33]))),IF(AND(RC[37]=RC[35],RC[38]=RC[-1]),(RC[34]-RC[39]+1)-(RC[32]+RC[33]),RC[34]-(RC[32]+RC[33]))))<0,0,IF(AND(RC[40]=""retiro"",RC[37]=RC[35],RC[38]=RC[-1]),IF((IF(RC[36]>(CONCATENATE(RC[35],RC[-1],RC[34])-RC[32]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[34]-RC[39])-RC[32])+1,(RC[34]-RC[39]+1)+(RC[34]-(RC[32]+RC[33]))),IF(AND(RC[37]=RC[35],RC[38]=RC[-1]),(RC[34]-RC[39]+1)-(RC[32]+RC[33]),RC[34]-(RC[32]+RC[33])))-IFERROR(VLOOKUP(RC[-7],'AISLAMIENTO COVID'!C1:C7,7,FALSE),0)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-8],Metas!C25:C31,6,FALSE),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-2]<=0,0,ROUND((IFERROR(VLOOKUP(RC[-9],Metas!C25:C31,6,0)*(RC[-2]/RC[32]),""Sin Metas"")),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-10],'NPS-UMBRAL'!C23:C24,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-4]=0,0%,RC[-1]/RC[-2])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(VLOOKUP(RC[-1],'Tabla Var'!R11C9:R14C11,3,TRUE)=""lineal"",RC[-1],VLOOKUP(RC[-1],'Tabla Var'!R11C9:R14C11,3,TRUE))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*R2C16"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(RC[-1]*((SUM(RC[-7],IF(IFERROR(VLOOKUP(RC[-14],'NPS-UMBRAL'!C1:C10,10,FALSE),0)>8,IFERROR(VLOOKUP(RC[-14],'NPS-UMBRAL'!C1:C10,10,FALSE),0),0))/RC[27]))*RC[-9])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-15],Metas!C25:C31,7,FALSE),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-9]<=0,0,ROUND((IFERROR(VLOOKUP(RC[-16],Metas!C25:C31,7,FALSE)*(RC[-9]/RC[25]),""Sin Metas"")),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-17],MOVIL!C1:C22,22,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-11]=0,0%,RC[-1]/RC[-2])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(VLOOKUP(RC[-1],'Tabla Var'!R11C9:R14C11,3,TRUE)=""lineal"",RC[-1],VLOOKUP(RC[-1],'Tabla Var'!R11C9:R14C11,3,TRUE))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*R2C23"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(RC[-1]*((SUM(RC[-14],IF(IFERROR(VLOOKUP(RC[-21],'NPS-UMBRAL'!C1:C10,10,FALSE),0)>8,IFERROR(VLOOKUP(RC[-21],'NPS-UMBRAL'!C1:C10,10,FALSE),0),0))/RC[20]))*RC[-16])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-22],Metas!C25:C32,8,FALSE),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(VLOOKUP(RC[-23],'NPS-UMBRAL'!C1:C17,15,0),0)=""NO REGISTRA"",""NO REGISTRA ENCUESTAS"",IFERROR(VLOOKUP(RC[-23],'NPS-UMBRAL'!C1:C17,15,0),""NO REGISTRA ENCUESTAS""))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(RC[-17]=0,0%,IF(RC[-1]<0%,0%,RC[-1]/RC[-2])),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[30]>=100%,IF(VLOOKUP(RC[-1],'Tabla Var'!R2C5:R13C7,3,TRUE)=""lineal"",'ROL-APP'!RC[-1],VLOOKUP(RC[-1],'Tabla Var'!R2C5:R13C7,3,TRUE)),IF(VLOOKUP(RC[-1],'Tabla Var'!R2C5:R13C8,4,TRUE)=""lineal"",'ROL-APP'!RC[-1],VLOOKUP(RC[-1],'Tabla Var'!R2C5:R13C8,4,TRUE)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*R2C29"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(RC[-1]*((SUM(RC[-20],IF(IFERROR(VLOOKUP(RC[-27],'NPS-UMBRAL'!C1:C10,10,FALSE),0)>8,IFERROR(VLOOKUP(RC[-27],'NPS-UMBRAL'!C1:C10,10,FALSE),0),0))/RC[14]))*RC[-22])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-28],Metas!C25:C34,10,FALSE),""Sin Metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-29],'NPS-UMBRAL'!C27:C28,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(RC[-23]=0,0%,IF(RC[-1]<0%,0%,RC[-1]/RC[-2])),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(VLOOKUP(RC[-1],'Tabla Var'!R11C9:R14C11,3,TRUE)=""lineal"",RC[-1],VLOOKUP(RC[-1],'Tabla Var'!R11C9:R14C11,3,TRUE))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*R2C35"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(RC[-1]*((SUM(RC[-26],IF(IFERROR(VLOOKUP(RC[-33],'NPS-UMBRAL'!C1:C10,10,FALSE),0)>8,IFERROR(VLOOKUP(RC[-33],'NPS-UMBRAL'!C1:C10,10,FALSE),0),0))/RC[8]))*RC[-28])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-8]+RC[-14]+RC[-21]+RC[-2]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(((RC[-1]*RC[-30])*((SUM(RC[-28],IF(IFERROR(VLOOKUP(RC[-35],'NPS-UMBRAL'!C1:C10,10,FALSE),0)>8,IFERROR(VLOOKUP(RC[-35],'NPS-UMBRAL'!C1:C10,10,FALSE),0),0))/RC[6]))),RC[16])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[12]>0,RC[-1]>0,RC[-1]<RC[12]),0,IF(AND(RC[12]>0,RC[-1]>0,RC[-1]>RC[12]),RC[-1]-RC[12],RC[-1]))"
    ActiveCell.Offset(0, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-39],'Ausentismos-Vaca-Umb'!C[-41]:C[-35],7,0),0)+IFERROR(IF(VLOOKUP(RC[-39],'NPS-UMBRAL'!C1:C11,10,0)>8,VLOOKUP(RC[-39],'NPS-UMBRAL'!C1:C11,10,0),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-40],'Ausentismos-Vaca-Umb'!C[-32]:C[-28],5,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C9,Meses!C[-43]:C[-41],3,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C9,Meses!C[-44]:C[-41],4,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-43],HC!C3:C6,4,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-1],1,4)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-2],5,2)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-3],7,2)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-47],'Ausentismos-Vaca-Umb'!C[-49]:C[-45],5,0),""S/N"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-48],Garantizado!C3:C5,3,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-49],HC!C3:C20,18,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-50],HC!C3:C12,10,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-44]=0,0,IFERROR(VLOOKUP(RC[-51],'DESARROLLO+PROYECTOS'!C3:C34,32,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-28],RC[-34],RC[-41],RC[-22])"
    ActiveCell.Offset(0, 1).Range("A1").Select
 ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-53],Metas!C25:C35,11,FALSE),""Sin metas"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "='Tabla Var'!R16C2"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-2]"
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlToLeft).Select
Next i
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
marcador:
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
End Sub


Sub FORMATITOS_7()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("ROL-APP").Select
Range("a1").Select
vari_9 = Range("a1").Value
Range("A4").Select
For i = 1 To vari_9

 ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[2],HC!C3:C13,11,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],HC!C3:C4,2,0)"
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlToLeft).Select
Next i

MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation

Call davanti


Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

Windows("PLAN_LIQ.xlsm").Activate
MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA ABRIR EL ARCHIVO Y EJECUTE EL BOTÓN GENERAR ARCHIVOS DE LIQUIDACIÓN", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.Quit

End Sub
Sub FORMATITOS_9()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("ASESORES TMK").Select
If Range("A1") > 1 Then

Sheets("ASESORES TMK").Select
Range("F4").Select
ActiveCell.Range("A1:B1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1:B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(-2, 0).Range("A1:B1").Select

Range("J4").Select

Else: Range("F4:G4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C3").Select
End If
MsgBox "LISTO EL PEGUE", vbInformation

Range("a1").Select
vari_9 = Range("a1").Value
Range("A4").Select
For i = 1 To vari_9

 ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[2],HC!C3:C13,11,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],HC!C3:C4,2,0)"
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlToLeft).Select
Next i

MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation


Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic



End Sub


Sub booo()
Sheets("Liquidacion_PART_TIME").Select
va_6 = Range("a1").Value
If va_6 = 0 Then
Sheets("Liquidacion_PART_TIME").Visible = False
Else
End If
Sheets("ROL-APP").Select
apenas = Range("a1").Value
If apenas = 0 Then
Sheets("ROL-APP").Visible = False
Else
End If
End Sub

Sub CAPRI()
Application.ScreenUpdating = False
Application.DisplayAlerts = False



Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Dim variable_1 As Object




Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\FIJAS SOLO CAV")


xiaomi = "FIJAS SOLO CAV"
    ss = xiaomi & ".xlsx"



Windows(ss).Activate

Sheets("Hoja1").Select
Range("A1").Select

Set variable_1 = Sheets("Hoja1").Range("AK:AK").Find(What:="MHBO", LookIn:=xlValues, LookAt:=xlWhole)

If Not variable_1 Is Nothing Then
Sheets("Hoja de pivote").Select
ActiveSheet.PivotTables("Tabla dinámica13").PivotSelect "", xlDataAndLabel, True
Selection.Copy


Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("AJ50").Select
ActiveSheet.Paste
 Application.CutCopyMode = False

Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=False

MsgBox "listo quinto MHBO", vbInformation
End If
 If variable_1 Is Nothing Then
 
 Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("AJ6") = "NO SE REGISTRAN VENTAS MHBO"
Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=False


MsgBox "listo quinto MHBO", vbInformation
End If
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub


Sub davanti()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

Sheets("Liquidacion").Select
Range("b1").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C50,""80,00%"")"
froome = Range("b1").Value
If froome > 0 Then

Range("AX3").AutoFilter Field:=50, Criteria1:="80,00%"
Range("AX1").Select
Selection.End(xlDown).Select

ActiveCell.FormulaR1C1 = _
        "=IF(RC[-40]=0,0%,IFERROR(IF(RC[-1]<=0%,0%,RC[-1]/RC[-2]),0))+0.000001"

ActiveCell.Offset(0, 0).Select
Selection.Copy
Range(Selection, Selection.End(xlDown)).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("AX1").Select
ActiveSheet.ShowAllData
Range("B1").ClearContents
MsgBox "validación del 80%, realizada", vbInformation

Else
Range("B1").ClearContents
MsgBox "validación del 80%, no se encontraron datos - ok", vbInformation

End If
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
End Sub
Sub uomo()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

Sheets("CEDULAS PILOTO").Select
Range("A2").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents
Range("D1").ClearContents
Range("A2").Select
Sheets("AISLAMIENTO COVID").Select
Range("A2").Select
rakuten = Selection.End(xlDown).Row
Sheets("AISLAMIENTO COVID").Range("A2:I" & rakuten).Select
Selection.ClearContents
Range("a1").Select

'Sheets("INCENTIVO").Select
'mana = Range("a1").Select
'If mana = "NO HAY INCENTIVO PARA ESTE MES" Then
'Range("A1").ClearContents
'Range("a1").Select
'Else
'Range("A1").Select
'Range(Selection, Selection.End(xlToRight)).Select
'Range(Selection, Selection.End(xlDown)).Select
'Selection.ClearContents
'Range("a1").Select
'End If

Sheets("HC").Select
Range("A1").Select
nike = Selection.End(xlDown).Row
Sheets("HC").Range("A2:AF" & nike).Select
Selection.ClearContents
Range("AC1").ClearContents
Columns("Z:AB").Select
Selection.Delete
Range("A1").Select

Sheets("Metas").Select
Range("A1").Select
puma = Selection.End(xlDown).Row
Sheets("Metas").Range("A3:AZ" & puma).Select
Selection.ClearContents
Range("X1").ClearContents
Range("A1").Select

Sheets("OTRAS GESTIONES").Select
Range("A3").Select
AKT = Selection.End(xlDown).Row
Sheets("OTRAS GESTIONES").Range("A4:AC" & AKT).Select
Selection.ClearContents
Range("A1, G2").ClearContents
Range("A1").Select


Sheets("Consultores_Inspira").Select
Range("A3").Select
AKT = Selection.End(xlDown).Row
Sheets("Consultores_Inspira").Range("A4:R" & AKT).Select
Selection.ClearContents
Range("A1, G2").ClearContents
Range("A1").Select



Sheets("ROL-APP").Select
Range("A3").Select
AKTT = Selection.End(xlDown).Row
Sheets("ROL-APP").Range("A4:BJ" & AKTT).Select
Selection.ClearContents
Range("A1, I2").ClearContents
Range("A1").Select
Range("F:G").Delete
Columns(6).EntireColumn.Insert
Columns(6).EntireColumn.Insert
Range("F3") = "NOMBRE DEL CAV"
 Range("G3") = "CÓDIGO DEL CAV"


Sheets("ASESORES TMK").Select
Range("A3").Select
AKTTO = Selection.End(xlDown).Row
Sheets("ASESORES TMK").Range("A4:AJ" & AKTTO).Select
Selection.ClearContents
Range("A1, I2").ClearContents
Range("A1").Select
Range("F:G").Delete
Columns(6).EntireColumn.Insert
Columns(6).EntireColumn.Insert
Range("F3") = "NOMBRE DEL CAV"
 Range("F3:G3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("F3:G3").Select
     Columns("F:G").Select
    Selection.ColumnWidth = 15
    Range("G4").Select

 
Sheets("Liquidacion").Select
Range("A3").Select
KOM = Selection.End(xlDown).Row
Sheets("Liquidacion").Range("A4:CL" & KOM).Select
Selection.ClearContents
Range("A1, I2").ClearContents
Range("F:G").Delete
Columns(6).EntireColumn.Insert
Columns(6).EntireColumn.Insert
Range("F3") = "NOMBRE DEL CAV"
Range("G3") = "CÓDIGO DEL CAV"
Range("A3").Select
'++
Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Select
Range("A3").Select
AFRICA = Selection.End(xlDown).Row
Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Range("A4:BZ" & AFRICA).Select
Selection.ClearContents
Range("A1, I2").ClearContents
Range("F:G").Delete
Columns(6).EntireColumn.Insert
Columns(6).EntireColumn.Insert
Range("F3") = "NOMBRE DEL CAV"
Range("G3") = "CÓDIGO DEL CAV"
Range("A3").Select
 '++
Sheets("Liquidacion_PART_TIME").Select
dell = Range("A1").Value
If dell > 0 Then
Range("A3").Select
banco = Selection.End(xlDown).Row
Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Range("A4:BY" & banco).Select
Selection.ClearContents
Range("A1, I2").ClearContents
Range("F:G").Delete
Columns(6).EntireColumn.Insert
Columns(6).EntireColumn.Insert
Range("F3") = "NOMBRE DEL CAV"
Range("G3") = "CÓDIGO DEL CAV"
Range("A3").Select
Else
Range("A1, I2").ClearContents
End If

Sheets("Resumen_plan_power").Select
Range("A6").Select
bici = Selection.End(xlDown).Row
Sheets("Resumen_plan_power").Range("A6:Q" & bici).Select
Selection.ClearContents
Range("Y1").ClearContents
Range("A6").Select


Sheets("Resumen Puntos").Select
Range("A6").Select
bicit = Selection.End(xlDown).Row
Sheets("Resumen Puntos").Range("A6:DB" & bicit).Select
Selection.ClearContents
Range("CE1").ClearContents
Range("A6").Select
Range("CP1").ClearContents

Sheets("VENTAS_OTRAS GESTIONES").Select
Range("A6").Select
icit = Selection.End(xlDown).Row
Sheets("VENTAS_OTRAS GESTIONES").Range("A6:CM" & icit).Select
Selection.ClearContents
Range("CO1, CP1").ClearContents
Range("A6").Select

Sheets("DESARROLLO+PROYECTOS").Select
Range("C2").Select
rikarena = Selection.End(xlDown).Row
Sheets("DESARROLLO+PROYECTOS").Range("C3:AN" & rikarena).Select
Selection.ClearContents
Range("AS1, G1").ClearContents
Range("A6").Select

Sheets("CUMPLIMIENTOS").Select
Range("A1").Select
rikar = Selection.End(xlDown).Row
Sheets("CUMPLIMIENTOS").Range("A2:F" & rikar).Select
Selection.ClearContents
Range("A6").Select

Sheets("Fuente Hogares").Select
Range("A:B").Select
Selection.ClearContents
Range("A2") = "ALTAS HOGAR"

Range("D:E").Select
Selection.ClearContents
Range("D2") = "CANALES ADULTOS"

Range("G:H").Select
Selection.ClearContents
Range("G2") = "TOTAL TECNOLOGIA"

Range("L:M").Select
Selection.ClearContents
Range("L2") = "CANALES PREMIUM CINE"

Range("P:Q").Select
Selection.ClearContents
Range("P5") = "VELOCIDAD 2"

Range("S:T").Select
Selection.ClearContents
Range("S5") = "VELOCIDAD 3"


Range("V:W").Select
Selection.ClearContents
Range("V5") = "VELOCIDAD 4"


Range("Z:AA").Select
Selection.ClearContents
Range("Z5") = "INGRESOS CAMBIO SERVICIOS FIJOS"

Range("AC:AD").Select
Selection.ClearContents
Range("AC5") = "INGRESOS MOVIL CAMBIOS DE PLAN MOVIL"

Range("AG:AH").Select
Selection.ClearContents
Range("AG3") = "CONVERGENCIA"

Range("AJ:AK").Select
Selection.ClearContents
Range("AJ5") = "MINIPACK HBO - FOX"

Range("AN:AO").Select
Selection.ClearContents
Range("AN3") = "NO RETENIDOS MOVIL (PLAN PAR - DESACTIV)"

Range("AQ:AR").Select
Selection.ClearContents
Range("AQ3") = "NO RETENIDOS SERVICIOS FIJOS"

Range("AT:AU").Select
Selection.ClearContents
Range("AT1") = "ULTRA WIFI - WIFI MESH"

Range("AW:AX").Select
Selection.ClearContents
Range("AW2") = "ARRIENDO - WIFI MESH"

Range("AZ:BA").Select
Selection.ClearContents
Range("AZ2") = "CANALES ADICIONALES"

Range("BC:BD").Select
Selection.ClearContents
Range("BC2") = "CANALES DEPORTIVOS"

Range("BE:BF").Select
Selection.ClearContents
Range("BE1") = "TECNOLOGÍA DE CONTADO"

Range("BH:BI").Select
Selection.ClearContents
Range("BH1") = "TECNOLOGÍA FINANCIADA"


Range("BL:BM").Select
Selection.ClearContents
Range("BL1") = "T- RESUELVE"
Range("A1").Select

Range("BO:BP").Select
Selection.ClearContents
Range("BO1") = "NETFLIX - SD-PREMIUM"
Range("A1").Select

Range("BS:BT").Select
Selection.ClearContents
Range("BS1") = "NETFLIX BASICO"
Range("A1").Select

Range("BZ:CA").Select
Selection.ClearContents
Range("CD:CE").Select
Selection.ClearContents

Range("BV5").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents
Range("BV5").Select

Sheets("Desarrollo").Select
Range("A1").Select
rik = Selection.End(xlDown).Row
Sheets("Desarrollo").Range("A2:C" & rik).Select
Selection.ClearContents
Range("A1").Select

Sheets("Proyectos").Select
Range("A1").Select
riki = Selection.End(xlDown).Row
Sheets("Proyectos").Range("A3:H" & riki).Select
Selection.ClearContents
Range("A1").Select


Sheets("METAS TMK").Select
Range("A1").Select
raketika = Selection.End(xlDown).Row
Sheets("METAS TMK").Range("A2:K" & raketika).Select
Selection.ClearContents
Range("A1").Select

Sheets("NPS-UMBRAL").Select
Range("A1").Select
CRACK = Selection.End(xlDown).Row
Sheets("NPS-UMBRAL").Range("A2:BD" & CRACK).Select
Selection.ClearContents
Range("P1").ClearContents
Range("A1").Select

Sheets("Ausentismos-Vaca-Umb").Select
Range("A1").Select
MIKE = Selection.End(xlDown).Row
Sheets("Ausentismos-Vaca-Umb").Range("A2:J" & MIKE).Select
Selection.ClearContents

Range("K1").Select
WEST = Selection.End(xlDown).Row
Sheets("Ausentismos-Vaca-Umb").Range("K2:AK" & WEST).Select
Selection.ClearContents
Range("AN1:AO1").ClearContents
Range("A1").Select
Sheets("Garantizado").Select
Range("A1").Select
SAKURA = Selection.End(xlDown).Row
Sheets("Garantizado").Range("A2:F" & SAKURA).Select
Selection.ClearContents
Range("H1").ClearContents
Range("A1").Select

'Sheets("MOVIL").Select
'
'ActiveSheet.PivotTables("Tabla dinámica18").PivotFields("CANAL").Orientation = _
'        xlHidden
'Range("A5").Select
'Range(Selection, Selection.End(xlToRight)).Select
'Range(Selection, Selection.End(xlDown)).Select
'Selection.ClearContents
'Range("A5").Select

Sheets("Tabla Var").Select
Range("U1").ClearContents
Range("U1").Select
MsgBox "listo primera fase", vbInformation
Call golovin


MsgBox "BORRADO COMPLETO, PLANTILLA", vbInformation
MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA A ABRIR EL ARCHIVO E INICIE LA LIQUIDACIÓN", vbInformation

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
ActiveWorkbook.Close SaveChanges:=True
Application.Quit
End Sub


Sub SAGAN()
Application.ScreenUpdating = False
Application.DisplayAlerts = False



Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Dim variable_1 As Object




Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\FIJAS SOLO CAV")


xiaomi = "FIJAS SOLO CAV"
    ss = xiaomi & ".xlsx"



Windows(ss).Activate

Sheets("Hoja1").Select
Range("A1").Select

Set variable_1 = Sheets("Hoja1").Range("AK:AK").Find(What:="MFOX", LookIn:=xlValues, LookAt:=xlWhole)

If Not variable_1 Is Nothing Then
Sheets("Hoja de pivote").Select
ActiveSheet.PivotTables("Tabla dinámica12").PivotSelect "", xlDataAndLabel, True
Selection.Copy


Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("AJ7").Select
ActiveSheet.Paste
 Application.CutCopyMode = False

Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=False

MsgBox "listo quinto MFOX", vbInformation
End If
 If variable_1 Is Nothing Then
 
 Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("AJ7") = "NO SE REGISTRAN VENTAS MFOX"
Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=False


MsgBox "listo quinto MFOX", vbInformation
End If
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Sub finestre()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

On Error Resume Next

Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String


Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\FIJAS SOLO CAV")


xiaomi = "FIJAS SOLO CAV"
    ss = xiaomi & ".xlsx"

Windows(ss).Activate


Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)

Worksheets.Add
ActiveSheet.Name = "Hoja de pivote2"

Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("A1"))

With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("TABLA")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Tipo venta")
        .Orientation = xlPageField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("tipo_contrato")
        .Orientation = xlPageField
        .Position = 3
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica1").PivotFields("Renta actual"), "Suma de Renta actual", xlSum
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("tipo_contrato"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("tipo_contrato"). _
        CurrentPage = "CONTADO"
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Tipo venta"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Tipo venta"). _
        CurrentPage = "T"
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("TABLA").ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("TABLA").CurrentPage = _
        "SRVADC"


With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Producto Liquid")
        .Orientation = xlPageField
        .Position = 3
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Producto Liquid"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Producto Liquid")
        .PivotItems("UWARR").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Producto Liquid"). _
        EnableMultiplePageItems = True



With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("MARCA SIN TURNO")
        .Orientation = xlPageField
        .Position = 4
    End With
     ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("MARCA SIN TURNO"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("MARCA SIN TURNO")
        .PivotItems("X").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("MARCA SIN TURNO"). _
        EnableMultiplePageItems = True


ActiveSheet.PivotTables("Tabla dinámica1").PivotSelect "", xlDataAndLabel, True
Selection.Copy

Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("BE4").Select
ActiveSheet.Paste
 Application.CutCopyMode = False

Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=True
Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA ABRIR EL ARCHIVO Y EJECUTE EL BOTON SÉPTIMA PARTE", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.Quit
End Sub
Sub tmk()


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("Tabla Var").Select
va_7 = Range("U1").Value
Sheets("ASESORES TMK").Select
Range("I2") = va_7
Range("a1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C[2])-2"
vari_9 = Range("a1").Value
Range("c4").Select

For i = 1 To vari_9

ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],HC!C3:C5,3,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],HC!C3:C7,5,0)"
    ActiveCell.Offset(0, 1).Range("A1:B1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-3],'METAS TMK'!C2:C4,3,0),""Sin Oficina"")"
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(VLOOKUP(RC[-5],HC!C3:C9,7,0),0)>0,VLOOKUP(RC[-5],'DESARROLLO+PROYECTOS'!C3:C9,7,FALSE),IFERROR(VLOOKUP(RC[-5],HC!C3:C9,7,0),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C9,Meses!C[-8]:C[-7],2,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((IF((IF(AND(RC[20]=""retiro"",RC[17]=RC[15],RC[18]=RC[-1]),IF((IF(RC[16]>(CONCATENATE(RC[15],RC[-1],RC[14])-RC[12]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[14]-RC[19])-RC[12])+1,(RC[14]-RC[19]+1)+(RC[14]-(RC[12]+RC[13]))),IF(AND(RC[17]=RC[15],RC[18]=RC[-1]),(RC[14]-RC[19]+1)-(RC[12]+RC[13]),RC[14]-(RC[12]+RC[13]))))<0,0,IF(AND(RC[20]=""retiro"",RC[17]=RC[15],RC[18]=RC[-1]),IF((IF(RC[16]>(CONCATENATE(RC[15],RC[-1],RC[14])-RC[12]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[14]-RC[19])-RC[12])+1,(RC[14]-RC[19]+1)+(RC[14]-(RC[12]+RC[13]))),IF(AND(RC[17]=RC[15],RC[18]=RC[-1]),(RC[14]-RC[19]+1)-(RC[12]+RC[13]),RC[14]-(RC[12]+RC[13])))-IFERROR(VLOOKUP(RC[-7],'AISLAMIENTO COVID'!C1:C7,7,FALSE),0)))<0,0," & _
        "IF((IF(AND(RC[20]=""retiro"",RC[17]=RC[15],RC[18]=RC[-1]),IF((IF(RC[16]>(CONCATENATE(RC[15],RC[-1],RC[14])-RC[12]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[14]-RC[19])-RC[12])+1,(RC[14]-RC[19]+1)+(RC[14]-(RC[12]+RC[13]))),IF(AND(RC[17]=RC[15],RC[18]=RC[-1]),(RC[14]-RC[19]+1)-(RC[12]+RC[13]),RC[14]-(RC[12]+RC[13]))))<0,0,IF(AND(RC[20]=""retiro"",RC[17]=RC[15],RC[18]=RC[-1]),IF((IF(RC[16]>(CONCATENATE(RC[15],RC[-1],RC[14])-RC[12]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[14]-RC[19])-RC[12])+1,(RC[14]-RC[19]+1)+(RC[14]-(RC[12]+RC[13]))),IF(AND(RC[17]=RC[15],RC[18]=RC[-1]),(RC[14]-RC[19]+1)-(RC[12]+RC[13]),RC[14]-(RC[12]+RC[13])))-IFERROR(VLOOKUP(RC[-7],'AISLAMIENTO COVID'!C1:C7,7,FALSE),0)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-8],'METAS TMK'!C2:C11,10,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=ROUND((VLOOKUP(RC[-9],'METAS TMK'!C2:C11,10,FALSE))*(RC[-2]/RC[12]),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(IFERROR(VLOOKUP(RC[-10],MOVIL!C1:C26,5,FALSE),0),IFERROR(VLOOKUP('ASESORES TMK'!RC[-10],MOVIL!C1:C26,9,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-4]=0,0,IF(AND(RC[-2]=0,RC[-1]>=1),100%,(RC[-1]/RC[-2])))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(VLOOKUP(RC[-1],'Tabla Var'!R2C1:R13C3,3,TRUE)=""lineal"",RC[-1],VLOOKUP(RC[-1],'Tabla Var'!R2C1:R13C3,3,TRUE))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*R2C16"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(RC[-1]*(SUM(RC[-7],IF(IFERROR(VLOOKUP(RC[-14],'NPS-UMBRAL'!C1:C10,10,FALSE),0)>8,IFERROR(VLOOKUP(RC[-14],'NPS-UMBRAL'!C1:C10,10,FALSE),0),0))/RC[7]))*RC[-9]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(((RC[-2]*RC[-10])*(SUM(RC[-8],IF(IFERROR(VLOOKUP(RC[-15],'NPS-UMBRAL'!C1:C10,10,FALSE),0)>8,IFERROR(VLOOKUP(RC[-15],'NPS-UMBRAL'!C1:C10,10,FALSE),0),0))/RC[6])),RC[16])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[12]>0,RC[-1]>0,RC[-1]<RC[12]),0,IF(AND(RC[12]>0,RC[-1]>0,RC[-1]>RC[12]),RC[-1]-RC[12],RC[-1]))"
    ActiveCell.Offset(0, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-19],'Ausentismos-Vaca-Umb'!C[-21]:C[-15],7,0),0)+IFERROR(IF(VLOOKUP(RC[-19],'NPS-UMBRAL'!C1:C11,10,0)>8,VLOOKUP(RC[-19],'NPS-UMBRAL'!C1:C11,10,0),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-20],'Ausentismos-Vaca-Umb'!C[-12]:C[-8],5,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C9,Meses!C[-23]:C[-21],3,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C9,Meses!C[-24]:C[-21],4,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-23],HC!C3:C6,4,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-1],1,4)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-2],5,2)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-3],7,2)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-27],'Ausentismos-Vaca-Umb'!C[-29]:C[-25],5,0),""S/N"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-28],Garantizado!C3:C5,3,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-29],HC!C3:C20,18,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-30],HC!C3:C12,10,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-31],'DESARROLLO+PROYECTOS'!C3:C34,32,FALSE),0)"
ActiveCell.Offset(0, 1).Range("A1").Select
ActiveCell.FormulaR1C1 = "=RC[-21]"
ActiveCell.Offset(1, 0).Select
Selection.End(xlToLeft).Select

    
 Next i
 
 
 Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
End Sub
Sub otricas()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("Tabla Var").Select
va_7 = Range("U1").Value
Sheets("OTRAS GESTIONES").Select
Range("G2") = va_7
Range("a1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C[2])-2"
vari_9 = Range("a1").Value
Range("c4").Select

For i = 1 To vari_9

ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],HC!C3:C5,3,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],HC!C3:C7,5,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(VLOOKUP(RC[-3],HC!C3:C9,7,0),0)>0,VLOOKUP(RC[-3],'DESARROLLO+PROYECTOS'!C3:C9,7,FALSE),IFERROR(VLOOKUP(RC[-3],HC!C3:C9,7,0),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C7,Meses!C[-6]:C[-5],2,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((IF((IF(AND(RC[15]=""retiro"",RC[12]=RC[10],RC[13]=RC[-1]),IF((IF(RC[11]>(CONCATENATE(RC[10],RC[-1],RC[9])-RC[7]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[9]-RC[14])-RC[7])+1,(RC[9]-RC[14]+1)+(RC[9]-(RC[7]+RC[8]))),IF(AND(RC[12]=RC[10],RC[13]=RC[-1]),(RC[9]-RC[14]+1)-(RC[7]+RC[8]),RC[9]-(RC[7]+RC[8]))))<0,0,IF(AND(RC[15]=""retiro"",RC[12]=RC[10],RC[13]=RC[-1]),IF((IF(RC[11]>(CONCATENATE(RC[10],RC[-1],RC[9])-RC[7]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[9]-RC[14])-RC[7])+1,(RC[9]-RC[14]+1)+(RC[9]-(RC[7]+RC[8]))),IF(AND(RC[12]=RC[10],RC[13]=RC[-1]),(RC[9]-RC[14]+1)-(RC[7]+RC[8]),RC[9]-(RC[7]+RC[8])))-IFERROR(VLOOKUP(RC[-5],'AISLAMIENTO COVID'!C1:C7,7,FALSE),0)))<0,0,IF((IF(AND(RC[15]=""retiro""," & _
        "RC[12]=RC[10],RC[13]=RC[-1]),IF((IF(RC[11]>(CONCATENATE(RC[10],RC[-1],RC[9])-RC[7]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[9]-RC[14])-RC[7])+1,(RC[9]-RC[14]+1)+(RC[9]-(RC[7]+RC[8]))),IF(AND(RC[12]=RC[10],RC[13]=RC[-1]),(RC[9]-RC[14]+1)-(RC[7]+RC[8]),RC[9]-(RC[7]+RC[8]))))<0,0,IF(AND(RC[15]=""retiro"",RC[12]=RC[10],RC[13]=RC[-1]),IF((IF(RC[11]>(CONCATENATE(RC[10],RC[-1],RC[9])-RC[7]),""retiro bien"",""retiro mal""))=""retiro mal"",((RC[9]-RC[14])-RC[7])+1,(RC[9]-RC[14]+1)+(RC[9]-(RC[7]+RC[8]))),IF(AND(RC[12]=RC[10],RC[13]=RC[-1]),(RC[9]-RC[14]+1)-(RC[7]+RC[8]),RC[9]-(RC[7]+RC[8])))-IFERROR(VLOOKUP(RC[-5],'AISLAMIENTO COVID'!C1:C7,7,FALSE),0)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
          ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-3]=0,RC[-1]=0),0,VLOOKUP(RC[-6],'VENTAS_OTRAS GESTIONES'!C1:C99,99,FALSE))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "80%"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=((RC[-5]*80%)*(RC[-3]/RC[6]))+RC[-2]+RC[16]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[12]>0,RC[-1]>0,RC[-1]<RC[12]),0,IF(AND(RC[12]>0,RC[-1]>0,RC[-1]>RC[12]),RC[-1]-RC[12],RC[-1]))"
    ActiveCell.Offset(0, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-12],'Ausentismos-Vaca-Umb'!C[-14]:C[-8],7,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-13],'Ausentismos-Vaca-Umb'!C[-5]:C[-1],5,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C7,Meses!C[-16]:C[-14],3,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C7,Meses!C[-17]:C[-14],4,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-16],HC!C3:C6,4,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-1],1,4)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-2],5,2)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-3],7,2)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-20],'Ausentismos-Vaca-Umb'!C[-22]:C[-18],5,0),""S/N"")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-21],Garantizado!C3:C5,3,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-22],HC!C3:C20,18,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-23],HC!C3:C12,10,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-19]=0,0,IFERROR(VLOOKUP(RC[-24],'DESARROLLO+PROYECTOS'!C3:C34,32,FALSE),0))"
ActiveCell.Offset(1, 0).Select
Selection.End(xlToLeft).Select

    
 Next i
 
 
 Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
End Sub

Sub power()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("Resumen Puntos").Select
Range("A6").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("Resumen_plan_power").Select
Range("A6").Select
ActiveSheet.Paste
Range("Y1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C1)-2"
vari_9 = Range("Y1").Value
Range("A6").Select
For i = 1 To vari_9
ActiveCell.Offset(0, 1).Range("A1").Select
   ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-1],Liquidacion!C3:C20,16,FALSE),IFERROR(VLOOKUP(Resumen_plan_power!RC[-1],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C20,16,FALSE),""CÉDULA NO ENCONTRADA""))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-2],Liquidacion!C3:C26,22,FALSE),IFERROR(VLOOKUP(Resumen_plan_power!RC[-2],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C26,22,FALSE),""CÉDULA NO ENCONTRADA""))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-2]>=R3C2,RC[-1]>=R3C3),RC[1]*3000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-4],MOVIL!C1:C31,31,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-4]>=R3C2,RC[-3]>=R3C3),RC[1]*6000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6],MOVIL!C1:C32,32,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-7],MOVIL!C53:C55,3,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(IFERROR(VLOOKUP(RC[-8],'% CUM_4 MES'!C1:C16,16,FALSE),0)>=R3C2,IFERROR(VLOOKUP(RC[-8],'% CUM_4 MES'!C1:C20,20,FALSE),0)>=R3C3),RC[-1]*2000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-9],MOVIL!C53:C54,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(IFERROR(VLOOKUP(RC[-10],'% CUM_4 MES'!C1:C16,16,FALSE),0)>=R3C2,IFERROR(VLOOKUP(RC[-10],'% CUM_4 MES'!C1:C20,20,FALSE),0)>=R3C3),RC[-1]*5000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-11],MOVIL!C53:C56,4,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(IFERROR(VLOOKUP(RC[-12],'% CUM_4 MES'!C1:C16,16,FALSE),0)>=R3C2,IFERROR(VLOOKUP(RC[-12],'% CUM_4 MES'!C1:C20,20,FALSE),0)>=R3C3),RC[-1]*10000,0)"
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-2],RC[-4],RC[-6],RC[-9],RC[-11],RC[1])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-14]>=R3C2,RC[-13]>=R3C3),RC[1]*4000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-16],MOVIL!C1:C37,37,FALSE),0)"
    

ActiveCell.Offset(1, 0).Select
Selection.End(xlToLeft).Select

    
 Next i
 
 
 Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
End Sub

Sub despro()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("Tabla Var").Select
va_7 = Range("U1").Value
Sheets("DESARROLLO+PROYECTOS").Select
Range("G1") = va_7
Range("AS1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C3)-1"
vari_9 = Range("AS1").Value
Range("C3").Select

For i = 1 To vari_9
ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],HC!C3:C7,3,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],HC!C3:C7,5,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],HC!C3:C20,18,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=R1C7"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],HC!C3:C9,7,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC3,HC!C3:C9,7,0)*RC[1]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=100%-RC[3]-RC[14]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "PROYECTOS"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-9],Proyectos!C2:C4,3,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-1]=100%,RC[10]>0),RC[-1]-RC[10],RC[-1])"
    ActiveCell.Offset(0, 4).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-14],Proyectos!C2:C5,4,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<80%,RC[-1],IF(RC[-1]>=80%,RC[-1]+10%,0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-6]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-3]>=90%,RC[13]>=100%),(RC[13]-100%+RC[-2])*RC[-7],RC[-1])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-13]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "DESARROLLO"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R1C9=""SI"",IF(ISNUMBER(SEARCH(""Gerente"",RC[-18])),20%,IF(ISNUMBER(SEARCH(""Director"",RC[-18])),20%,10%)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]"
    ActiveCell.Offset(0, 4).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-25],Desarrollo!C1:C3,3,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(LOOKUP(RC[-1],'Tabla Var'!R2C31:R8C33,'Tabla Var'!R2C33:R8C33),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-6]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(COUNT(RC[6]:RC[8])>=2,RC[-3]>=90%,RC[9]>=100%),(RC[9]-100%+RC[-2])*RC[-7],RC[-1])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[-24]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND((IFERROR(VLOOKUP(RC[-30],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C71,69,FALSE),IFERROR(VLOOKUP('DESARROLLO+PROYECTOS'!RC[-30],Liquidacion!C3:C82,80,FALSE),IFERROR(VLOOKUP('DESARROLLO+PROYECTOS'!RC[-30],'ASESORES TMK'!C3:C35,33,FALSE),IFERROR(VLOOKUP('DESARROLLO+PROYECTOS'!RC[-30],'ROL-APP'!C3:C62,60,FALSE),""CÉDULA NO ENCONTRADA"")))))>=160%,(IFERROR(VLOOKUP(RC[-30],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C71,69,FALSE),IFERROR(VLOOKUP('DESARROLLO+PROYECTOS'!RC[-30],Liquidacion!C3:C82,80,FALSE),IFERROR(VLOOKUP('DESARROLLO+PROYECTOS'!RC[-30],'ASESORES TMK'!C3:C35,33,FALSE),IFERROR(VLOOKUP('DESARROLLO+PROYECTOS'!RC[-30],'ROL-APP'!C3:C62,60,FALSE),""CÉDULA NO ENCONTRADA"")))))<>""CÉDULA NO ENCONTRADA""),160%," & _
        "IFERROR(VLOOKUP(RC[-30],'Liquid Tiendas-Cvc-DENTRO CAV'!C3:C71,69,FALSE),IFERROR(VLOOKUP('DESARROLLO+PROYECTOS'!RC[-30],Liquidacion!C3:C82,80,FALSE),IFERROR(VLOOKUP('DESARROLLO+PROYECTOS'!RC[-30],'ASESORES TMK'!C3:C35,33,FALSE),IFERROR(VLOOKUP('DESARROLLO+PROYECTOS'!RC[-30],'ROL-APP'!C3:C62,60,FALSE),""CÉDULA NO ENCONTRADA"")))))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-13]"
    ActiveCell.Offset(0, 3).Select
      ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-34],CUMPLIMIENTOS!C1:C4,4,FALSE),"""")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-35],CUMPLIMIENTOS!C1:C5,5,FALSE),"""")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-36],CUMPLIMIENTOS!C1:C6,6,FALSE),"""")"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((IFERROR(AVERAGE(RC[-3]:RC[-1]),0))>=160%,160%,IFERROR(AVERAGE(RC[-3]:RC[-1]),0))"
    
ActiveCell.Offset(1, 0).Select
Selection.End(xlToLeft).Select

    
 Next i
 
 
 Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation

End Sub
Sub proyecto()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Dim archivos

Dim vari_2 As Excel.Workbook

Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\PROYECTOS\*.xlsx")
If archivos = "" Then
MsgBox "NO HAY NINGÚN ARCHIVO DE PROYECTOS CARGADO", vbCritical
Else
Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\PROYECTOS\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate
Sheets("Resultado").Select

vegas = Sheets("Resultado").Range("A" & Rows.Count).End(xlUp).Row
Sheets("Resultado").Range("A4:E" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Proyectos").Select
Range("A3").PasteSpecial xlPasteValues
Application.CutCopyMode = False
Range("A1").Select
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
End If
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "ok primera fase", vbInformation
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
End Sub

Sub desarrollo()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Dim archivos

Dim vari_2 As Excel.Workbook

Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\DESARROLLO\*.xlsx")
If archivos = "" Then
MsgBox "NO HAY NINGÚN ARCHIVO DE DESARROLLO CARGADO", vbCritical
Else
Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\DESARROLLO\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate
vegas = Sheets("VARIABLE").Range("A" & Rows.Count).End(xlUp).Row
Sheets("VARIABLE").Range("A2:B" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Desarrollo").Select
Range("A2").PasteSpecial xlPasteValues
Application.CutCopyMode = False
Windows(tt).Activate
Sheets("VARIABLE").Range("E2:E" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Range("C2").PasteSpecial xlPasteValues
Application.CutCopyMode = False
Range("A1").Select
Windows(tt).Activate
Sheets("VARIABLE").Select
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
End If
 Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA ABRIR EL ARCHIVO Y EJECUTE EL BOTÓN APLICAR FORMATOS", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.Quit

End Sub
Sub tmk_vendehumo()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos

Dim vari_2 As Excel.Workbook

Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\METAS\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\METAS\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate

Sheets("Fuente_Consultor Apoyo TMK").Select
vegas = Sheets("Fuente_Consultor Apoyo TMK").Range("A" & Rows.Count).End(xlUp).Row
Sheets("Fuente_Consultor Apoyo TMK").Range("A2:I" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("METAS TMK").Select
Range("A2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "Listo el pegue", vbInformation
Windows("PLAN_LIQ.xlsm").Activate
Sheets("METAS TMK").Select
Range("A1").Select
roglic = Selection.End(xlDown).Row - 1
Range("k2").Select
For i = 1 To roglic
ActiveCell.FormulaR1C1 = "=SUM(RC[-4]:RC[-3])"
ActiveCell.Offset(1, 0).Select
Next i
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "Listo primera fase", vbInformation
End Sub
Sub everest()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos

Dim vari_2 As Excel.Workbook

Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

Windows("PLAN_LIQ.xlsm").Activate
Sheets("HC").Select
Set rangodatos = Sheets("HC").UsedRange
    rangodatos.AutoFilter Field:=7, Criteria1:=Array("Consultor Integral Servicio A Clientes", "Asesor Servicio Al Cliente", "Consultor Integral Servicio A Clientes Sr", "Asesor Integral Servicio Al Cliente", "Consultor(a) Integral Servicio A Clientes", "Asesor(a) Servicio Al Cliente", "Consultor(a) Integral Servicio A Clientes Sr", "Consultor(a) Servicio Personalizado A Clientes", "Asesor(a) Integral Servicio Al Cliente"), _
        Operator:=xlFilterValues
Range("C1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("DESARROLLO+PROYECTOS").Select
Range("C3").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("C3").Select
Sheets("HC").Select
Range("A1").Select
ActiveSheet.ShowAllData
MsgBox "listo el pegue", vbInformation

archivos = Dir("D:\AUTOMATIZACION\DESARROLLO\*.xlsx")
If archivos = "" Then
MsgBox "NO HAY NINGÚN ARCHIVO DE DESARROLLO CARGADO", vbCritical
Else
Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\DESARROLLO\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate
Sheets("Cumplimientos").Select
Sheets("Cumplimientos").Select
vegas = Sheets("Cumplimientos").Range("A" & Rows.Count).End(xlUp).Row
Sheets("Cumplimientos").Range("A2:F" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("CUMPLIMIENTOS").Select
Range("A2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A2").Select
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
End If
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Sub agenditas()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos

Dim vari_2 As Excel.Workbook

Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\TRANSACCIONES_APP\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\TRANSACCIONES_APP\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate

Sheets("Aprovechamiento de Tráfico").Select
vegas = Sheets("Aprovechamiento de Tráfico").Range("A" & Rows.Count).End(xlUp).Row
Sheets("Aprovechamiento de Tráfico").Range("A2:A" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("NPS-UMBRAL").Select
Range("AA2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows(tt).Activate
Sheets("Aprovechamiento de Tráfico").Range("D2:D" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("NPS-UMBRAL").Select
Range("AB2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Range("A2").Select
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub
Sub coronita()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos

Dim vari_2 As Excel.Workbook

Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\COVID\*.xlsx")
If archivos = "" Then
MsgBox "NO HAY NINGÚN ARCHIVO DE COVID CARGADO", vbCritical
Else
Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\COVID\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate
Sheets(1).Select
vegas = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("A2:H" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("AISLAMIENTO COVID").Select
Range("A2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A2").Select
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
End If
MsgBox "LISTO SEGUNDA FASE", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True


End Sub
Sub finestre_2()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

On Error Resume Next

Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String


Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\FIJAS SOLO CAV")


xiaomi = "FIJAS SOLO CAV"
    ss = xiaomi & ".xlsx"

Windows(ss).Activate


Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)

Worksheets.Add
ActiveSheet.Name = "Hoja de pivote3"

Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("A1"))

With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("TABLA")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Tipo venta")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").AddDataField ActiveSheet.PivotTables _
        ("Tabla dinámica1").PivotFields("Renta actual"), "Suma de Renta actual", xlSum
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("tipo_contrato")
        .Orientation = xlPageField
        .Position = 2
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("TABLA").ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("TABLA").CurrentPage = _
        "SRVADC"
    ActiveSheet.PivotTables("Tabla dinámica1").PivotSelect _
        "'Tipo venta'[All] SRVADC", xlDataAndLabel, True
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("tipo_contrato"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("tipo_contrato"). _
        CurrentPage = "FINANCIADO"
    Range("B4").Select
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Tipo venta"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Tipo venta"). _
        CurrentPage = "T"
  
  With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Producto Liquid")
        .Orientation = xlPageField
        .Position = 3
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Producto Liquid"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Producto Liquid")
        .PivotItems("UWARR").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Producto Liquid"). _
        EnableMultiplePageItems = True
  
  
  With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("MARCA SIN TURNO")
        .Orientation = xlPageField
        .Position = 4
    End With
     ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("MARCA SIN TURNO"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("MARCA SIN TURNO")
        .PivotItems("X").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("MARCA SIN TURNO"). _
        EnableMultiplePageItems = True
   
  
ActiveSheet.PivotTables("Tabla dinámica1").PivotSelect "", xlDataAndLabel, True
Selection.Copy


Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("BH4").Select
ActiveSheet.Paste
 Application.CutCopyMode = False
Range("A1").Select


'++

Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=True
Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA ABRIR EL ARCHIVO Y EJECUTE EL BOTON SÉPTIMA PARTE", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.Quit


End Sub

Sub ineos()

'

'
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("OTRAS GESTIONES").Select
Range("C4").Select
Range(Selection, Selection.End(xlDown)).Copy
Sheets("VENTAS_OTRAS GESTIONES").Select
Range("A6").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A6").Select
MsgBox "listo primera fase", vbInformation
Range("CP1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C1)-1"
vari_9 = Range("CP1").Value
Range("A6").Select

For i = 1 To vari_9
'

ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "90%"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "90%"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "90%"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "90%"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5],MOVIL!C1:C26,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]>=R3C4,RC[-1]*23000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-7],MOVIL!C1:C26,3,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-5]>=R3C4,RC[-1]*17000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-9],MOVIL!C1:C26,4,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-7]>=R3C4,RC[-1]*7000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-11],MOVIL!C1:C6,6,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-10]>=R3C3,RC[-9]>=R3C4),RC[-1]*23000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-13],MOVIL!C1:C7,7,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-12]>=R3C3,RC[-11]>=R3C4),RC[-1]*17000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-15],MOVIL!C1:C8,8,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-14]>=R3C3,RC[-13]>=R3C4),RC[-1]*7000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-1],RC[-3],RC[-5],RC[-7],RC[-9],RC[-11])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(VLOOKUP(RC[-18],MOVIL!C1:C14,14,FALSE),0)>=50,0,IFERROR(VLOOKUP(RC[-18],MOVIL!C1:C26,11,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(VLOOKUP(RC[-19],MOVIL!C1:C14,14,FALSE),0)>=50,0,IFERROR(VLOOKUP(RC[-19],MOVIL!C1:C26,12,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(VLOOKUP(RC[-20],MOVIL!C1:C14,14,FALSE),0)>=50,0,IFERROR(VLOOKUP(RC[-20],MOVIL!C1:C26,13,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-19]>=R3C3,RC[-18]>=R3C4),RC[1]*23000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-21]>=R3C3,RC[-20]>=R3C4),RC[1]*17000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-5]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-23]>=R3C3,RC[-22]>=R3C4),RC[1]*7000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-6]"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-2],RC[-4],RC[-6],RC[3])"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-25]>=R3C4,IFERROR(VLOOKUP(RC[-28],MOVIL!C1:C35,35,FALSE),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-26]>=R3C4,IFERROR(VLOOKUP(RC[-29],MOVIL!C1:C34,34,FALSE),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-27]>=R3C4,IF(IFERROR(VLOOKUP(RC[-30],MOVIL!C1:C14,14,FALSE),0)>=50,IFERROR(VLOOKUP(RC[-30],MOVIL!C1:C36,36,FALSE),0),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]*0.9%"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-32],MOVIL!C1:C30,27,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]*0.9%"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-34],MOVIL!C1:C29,29,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]*0.7%"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=(IFERROR(VLOOKUP(RC[-36],MOVIL!C1:C28,28,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]*0.7%"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=(IFERROR(VLOOKUP(RC[-38],MOVIL!C1:C30,30,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-36]>=R3C4,IFERROR(VLOOKUP(RC[-39],'Fuente Hogares'!C12:C13,2,FALSE)*5000,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-40],'Fuente Hogares'!C12:C13,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-38]>=R3C4,IFERROR(VLOOKUP(RC[-41],'Fuente Hogares'!C4:C5,2,FALSE)*5000,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-42],'Fuente Hogares'!C4:C5,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-40]>=R3C4,RC[1]*5000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(SUM(IFERROR(VLOOKUP(RC[-44],MOVIL!C1:C22,17,FALSE),0),IFERROR(VLOOKUP(RC[-44],MOVIL!C1:C22,18,FALSE),0)))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-42]>=R3C4,IFERROR((IFERROR(VLOOKUP(RC[-45],MOVIL!C1:C26,19,FALSE),0))*5500,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR((IFERROR(VLOOKUP(RC[-46],MOVIL!C1:C26,19,FALSE),0)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-44]>=R3C4,IFERROR((IFERROR(VLOOKUP(RC[-47],MOVIL!C1:C26,20,FALSE),0))*11000,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR((IFERROR(VLOOKUP(RC[-48],MOVIL!C1:C26,20,FALSE),0)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-46]>=R3C4,IFERROR((IFERROR(VLOOKUP(RC[-49],MOVIL!C1:C26,21,FALSE),0))*22000,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR((IFERROR(VLOOKUP(RC[-50],MOVIL!C1:C26,21,FALSE),0)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-48]>=R3C4,IFERROR((VLOOKUP(IFERROR(VLOOKUP(RC[-51],MOVIL!C1:C26,22,FALSE),0),'Tabla Var'!R3C9:R7C11,3,TRUE))*(IFERROR(VLOOKUP(RC[-51],MOVIL!C1:C26,22,FALSE),0)),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-52],MOVIL!C1:C26,22,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-50]>=R3C4,(((IFERROR(VLOOKUP(RC[-53],MOVIL!C1:C26,23,FALSE),0)))*1700),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=(IFERROR(VLOOKUP(RC[-54],MOVIL!C1:C26,23,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-52]>=R3C4,(((IFERROR(VLOOKUP(RC[-55],MOVIL!C1:C26,24,FALSE),0)))*3900),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-56],MOVIL!C1:C26,24,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-55]>=R3C3,RC[-54]>=R3C4),RC[1]*11000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-58],'Fuente Hogares'!C78:C79,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-56]>=R3C4,RC[1]*11000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-60],'Fuente Hogares'!C1:C2,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(AND(RC[-59]>=R3C3,RC[-58]>=R3C4),RC[1]*11000,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-62],'Fuente Hogares'!C82:C83,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]*0.9%"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(IFERROR(VLOOKUP(RC[-64],'Fuente Hogares'!C57:C58,2,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]*0.7%"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=(IFERROR(VLOOKUP(RC[-66],'Fuente Hogares'!C60:C61,2,FALSE),0))"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-64]>=R3C4,IFERROR(VLOOKUP(RC[-67],'Fuente Hogares'!C16:C17,2,0)*5500,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-68],'Fuente Hogares'!C16:C17,2,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-66]>=R3C4,((IFERROR(VLOOKUP(RC[-69],'Fuente Hogares'!C19:C20,2,FALSE),0)*11000)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-70],'Fuente Hogares'!C19:C20,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-68]>=R3C4,((IFERROR(VLOOKUP(RC[-71],'Fuente Hogares'!C22:C23,2,FALSE),0)*22000)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-72],'Fuente Hogares'!C22:C23,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-70]>=R3C4,((IFERROR(VLOOKUP(RC[-73],'Fuente Hogares'!C36:C37,2,FALSE),0))*2000),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-74],'Fuente Hogares'!C36:C37,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-72]>=R3C4,((IFERROR(VLOOKUP(RC[-75],'Fuente Hogares'!C46:C47,2,FALSE),0)*5000)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-76],'Fuente Hogares'!C46:C47,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-74]>=R3C4,IFERROR(VLOOKUP(RC[-77],MOVIL!C1:C26,26,FALSE),0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-75]>=R3C4,IF(IFERROR(VLOOKUP(RC[-78],'Fuente Hogares'!C74:C75,2,FALSE),0)<=10,(IFERROR(VLOOKUP(RC[-78],'Fuente Hogares'!C74:C75,2,FALSE),0)*2000),(IFERROR(VLOOKUP(RC[-78],'Fuente Hogares'!C74:C75,2,FALSE),0)*3000)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-79],'Fuente Hogares'!C74:C75,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-77]>=R3C4,((IFERROR(VLOOKUP(RC[-80],'Fuente Hogares'!C55:C56,2,FALSE),0)*5000)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-81],'Fuente Hogares'!C55:C56,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-79]>=R3C4,((IFERROR(VLOOKUP(RC[-82],'Fuente Hogares'!C52:C53,2,FALSE),0)*2200)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-83],'Fuente Hogares'!C52:C53,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-81]>=R3C4,RC[1]*30%,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-85],MOVIL!C1:C33,33,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-83]>=R3C4,((IFERROR(VLOOKUP(RC[-86],'Fuente Hogares'!C64:C65,2,FALSE),0)*2200)),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-87],'Fuente Hogares'!C64:C65,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-85]>=R3C4,RC[1]*10%,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-89],MOVIL!C1:C38,38,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-87]>=R3C4,(IFERROR(VLOOKUP(RC[-90],'Fuente Hogares'!C71:C72,2,FALSE),0)*2200),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-91],'Fuente Hogares'!C71:C72,2,FALSE),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-89]>=R3C4,(IFERROR(VLOOKUP(RC[-92],'Fuente Hogares'!C67:C68,2,FALSE),0)*5000),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-93],'Fuente Hogares'!C67:C68,2,FALSE),0)"
    ActiveCell.Offset(0, 5).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(RC[-6],RC[-8],RC[-10],RC[-12],RC[-14],RC[-16],RC[-18],RC[-20],RC[-21],RC[-23],RC[-25],RC[-27],RC[-29],RC[-31],RC[-33],RC[-35],RC[-37],RC[-39],RC[-41],RC[-43],RC[-45],RC[-47],RC[-49],RC[-51],RC[-53],RC[-55],RC[-57],RC[-59],RC[-61],RC[-63],RC[-65],RC[-67],RC[-71]:RC[-69],RC[-81])"
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlToLeft).Select

Next i

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation




End Sub

Sub virus()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("AISLAMIENTO COVID").Select
Range("A1").Select
goretzka = Selection.End(xlDown).Row - 1
Range("I2").Select
For i = 1 To goretzka
ActiveCell.FormulaR1C1 = "=IF(RC[-2]<20,""SUMAR DÍAS"",""PAGAR EL 80%"")"
ActiveCell.Offset(1, 0).Select
Next i
MsgBox "LISTO PRIMERA FASE", vbInformation
Sheets("OTRAS GESTIONES").Select
joao = Range("A1").Value
Range("AB4").Select
For j = 1 To joao
ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-25],'AISLAMIENTO COVID'!C1:C9,9,FALSE),""NO APLICA"")"
ActiveCell.Offset(1, 0).Select
Next j
Range("A1").Select
MsgBox "LISTO SEGUNDA FASE", vbInformation

Sheets("ROL-APP").Select
anabe = Range("A1").Value
Range("BD4").Select
For p = 1 To anabe
ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-53],'AISLAMIENTO COVID'!C1:C9,9,FALSE),""NO APLICA"")"
ActiveCell.Offset(1, 0).Select
Next p
Range("A1").Select
MsgBox "LISTO TERCERA FASE", vbInformation

Sheets("ASESORES TMK").Select
anabel = Range("A1").Value
Range("AI4").Select
For q = 1 To anabel
ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-32],'AISLAMIENTO COVID'!C1:C9,9,FALSE),""NO APLICA"")"
ActiveCell.Offset(1, 0).Select
Next q
Range("A1").Select
MsgBox "LISTO CUARTA FASE", vbInformation

Sheets("Liquidacion").Select
colorado = Range("A1").Value
Range("CD4").Select
For r = 1 To colorado
ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-79],'AISLAMIENTO COVID'!C1:C9,9,FALSE),""NO APLICA"")"
ActiveCell.Offset(1, 0).Select
Next r
Range("A1").Select
MsgBox "LISTO QUINTA FASE", vbInformation

Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Select
busi = Range("A1").Value
Range("BS4").Select
For t = 1 To busi
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-68],'AISLAMIENTO COVID'!C1:C9,9,FALSE),""NO APLICA"")"
ActiveCell.Offset(1, 0).Select
Next t
Range("A1").Select
MsgBox "LISTO QUINTA FASE", vbInformation

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation

End Sub
Sub liberman()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

Sheets("OTRAS GESTIONES").Select
busi = Range("A1").Value
Range("O4").Select
For t = 1 To busi
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[13]=""SUMAR DÍAS"",SUM(IFERROR(VLOOKUP(RC[-12],'Ausentismos-Vaca-Umb'!C[-14]:C[-8],7,0),0)+IFERROR(IF(VLOOKUP(RC[-12],'NPS-UMBRAL'!C1:C11,10,0)>8,VLOOKUP(RC[-12],'NPS-UMBRAL'!C1:C11,10,0),0),0),VLOOKUP(RC[-12],'AISLAMIENTO COVID'!C1:C7,7,FALSE)),IFERROR(VLOOKUP(RC[-12],'Ausentismos-Vaca-Umb'!C[-14]:C[-8],7,0),0)+IFERROR(IF(VLOOKUP(RC[-12],'NPS-UMBRAL'!C1:C11,10,0)>8,VLOOKUP(RC[-12],'NPS-UMBRAL'!C1:C11,10,0),0),0))"
ActiveCell.Offset(1, 0).Select
Next t
Range("A1").Select
MsgBox "LISTO PRIMERA FASE", vbInformation

Sheets("ROL-APP").Select
basi = Range("A1").Value
Range("AQ4").Select
For x = 1 To basi
    
ActiveCell.FormulaR1C1 = _
        "=IF(RC[13]=""SUMAR DÍAS"",SUM(IFERROR(VLOOKUP(RC[-40],'Ausentismos-Vaca-Umb'!C[-42]:C[-36],7,0),0)+IFERROR(IF(VLOOKUP(RC[-40],'NPS-UMBRAL'!C1:C11,10,0)>8,VLOOKUP(RC[-40],'NPS-UMBRAL'!C1:C11,10,0),0),0),VLOOKUP(RC[-40],'AISLAMIENTO COVID'!C1:C7,7,FALSE)),IFERROR(VLOOKUP(RC[-40],'Ausentismos-Vaca-Umb'!C[-42]:C[-36],7,0),0)+IFERROR(IF(VLOOKUP(RC[-40],'NPS-UMBRAL'!C1:C11,10,0)>8,VLOOKUP(RC[-40],'NPS-UMBRAL'!C1:C11,10,0),0),0))"
ActiveCell.Offset(1, 0).Select
Next x
Range("A1").Select
MsgBox "LISTO SEGUNDA FASE", vbInformation

Sheets("ASESORES TMK").Select
blaasi = Range("A1").Value
Range("V4").Select
For b = 1 To blaasi
    
ActiveCell.FormulaR1C1 = _
        "=IF(RC[13]=""SUMAR DÍAS"",SUM(IFERROR(VLOOKUP(RC[-19],'Ausentismos-Vaca-Umb'!C[-21]:C[-15],7,0),0)+IFERROR(IF(VLOOKUP(RC[-19],'NPS-UMBRAL'!C1:C11,10,0)>8,VLOOKUP(RC[-19],'NPS-UMBRAL'!C1:C11,10,0),0),0),VLOOKUP(RC[-19],'AISLAMIENTO COVID'!C1:C7,7,FALSE)),IFERROR(VLOOKUP(RC[-19],'Ausentismos-Vaca-Umb'!C[-21]:C[-15],7,0),0)+IFERROR(IF(VLOOKUP(RC[-19],'NPS-UMBRAL'!C1:C11,10,0)>8,VLOOKUP(RC[-19],'NPS-UMBRAL'!C1:C11,10,0),0),0))"
ActiveCell.Offset(1, 0).Select
Next b
Range("A1").Select
MsgBox "LISTO TERCERA FASE", vbInformation

Sheets("Liquidacion").Select
blsi = Range("A1").Value
Range("BQ4").Select
For w = 1 To blsi
  ActiveCell.FormulaR1C1 = _
        "=IF(RC[13]=""SUMAR DÍAS"",SUM(IFERROR(VLOOKUP(RC[-66],'Ausentismos-Vaca-Umb'!C[-68]:C[-62],7,0),0)+IFERROR(IF(VLOOKUP(RC[-66],'NPS-UMBRAL'!C1:C11,10,0)>8,VLOOKUP(RC[-66],'NPS-UMBRAL'!C1:C11,10,0),0),0),VLOOKUP(RC[-66],'AISLAMIENTO COVID'!C1:C7,7,FALSE)),IFERROR(VLOOKUP(RC[-66],'Ausentismos-Vaca-Umb'!C[-68]:C[-62],7,0),0)+IFERROR(IF(VLOOKUP(RC[-66],'NPS-UMBRAL'!C1:C11,10,0)>8,VLOOKUP(RC[-66],'NPS-UMBRAL'!C1:C11,10,0),0),0))"

ActiveCell.Offset(1, 0).Select
Next w
Range("A1").Select
MsgBox "LISTO CUARTA FASE", vbInformation

Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Select
bls = Range("A1").Value
Range("BF4").Select
For u = 1 To bls
  ActiveCell.FormulaR1C1 = _
        "=IF(RC[13]=""SUMAR DÍAS"",SUM(IFERROR(VLOOKUP(RC[-55],'Ausentismos-Vaca-Umb'!C[-57]:C[-51],7,0),0)+IFERROR(IF(VLOOKUP(RC[-55],'NPS-UMBRAL'!C1:C11,10,0)>8,VLOOKUP(RC[-55],'NPS-UMBRAL'!C1:C11,10,0),0),0),VLOOKUP(RC[-55],'AISLAMIENTO COVID'!C1:C7,7,FALSE)),IFERROR(VLOOKUP(RC[-55],'Ausentismos-Vaca-Umb'!C[-57]:C[-51],7,0),0)+IFERROR(IF(VLOOKUP(RC[-55],'NPS-UMBRAL'!C1:C11,10,0)>8,VLOOKUP(RC[-55],'NPS-UMBRAL'!C1:C11,10,0),0),0))"

ActiveCell.Offset(1, 0).Select
Next u
Range("A1").Select
MsgBox "LISTO QUINTA FASE", vbInformation

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
 End Sub

Sub bayern()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual


Sheets("ROL-APP").Select
basi = Range("A1").Value
Range("AL4").Select
For x = 1 To basi
    
ActiveCell.FormulaR1C1 = _
        "=IF(RC[18]=""PAGAR EL 80%"",SUM(((RC[-30]*80%)*(RC[-28]/RC[7])),RC[17]),SUM(((RC[-1]*RC[-30])*(RC[-28]/RC[7])),RC[17]))"

ActiveCell.Offset(1, 0).Select
Next x
Range("A1").Select
MsgBox "LISTO PRIMERA FASE", vbInformation

Sheets("ASESORES TMK").Select
blaasi = Range("A1").Value
Range("R4").Select
For b = 1 To blaasi
    
ActiveCell.FormulaR1C1 = _
        "=IF(RC[17]=""PAGAR EL 80%"",SUM(((RC[-10]*80%)*(RC[-8]/RC[6])),RC[16]),SUM(((RC[-2]*RC[-10])*(RC[-8]/RC[6])),RC[16]))"
ActiveCell.Offset(1, 0).Select
Next b
Range("A1").Select
MsgBox "LISTO SEGUNDA FASE", vbInformation


Sheets("Liquidacion").Select
blsi = Range("A1").Value
Range("BK4").Select
For w = 1 To blsi
  ActiveCell.FormulaR1C1 = _
        "=IF(RC[19]=""PAGAR EL 80%"",SUM(((RC[-55]*80%)*(RC[-53]/RC[8])),RC[18]),SUM((IF((AND(RC[-26]=0,OR(RC[-19]=0,RC[-19]=""""))),0,SUM((RC[-3],RC[-29],RC[-1])))),RC[18]))"

ActiveCell.Offset(1, 0).Select
Next w
Range("A1").Select
MsgBox "LISTO TERCERA FASE", vbInformation

Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Select
bls = Range("A1").Value
Range("AZ4").Select
For u = 1 To bls
      ActiveCell.FormulaR1C1 = _
        "=IF(RC[19]=""PAGAR EL 80%"",SUM(((RC[-44]*80%)*(RC[-42]/RC[8])),RC[18]),SUM(RC[-3],RC[-18],RC[-1],RC[18]))"

ActiveCell.Offset(1, 0).Select
Next u
Range("A1").Select
MsgBox "LISTO CUARTA FASE", vbInformation

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation


End Sub

Sub golovin()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim vari_2 As Excel.Workbook
Dim vari_3 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\FIJAS SOLO CAV")
xiaomi = "FIJAS SOLO CAV"
ss = xiaomi & ".xlsx"
Sheets("Hoja de pivote4").Delete
Sheets("Hoja de pivote3").Delete
Sheets("Hoja de pivote2").Delete
Sheets("Hoja de pivote").Delete
Sheets("Hoja1").Select
Range("A1").Select
coll = Selection.End(xlDown).Row
Sheets("Hoja1").Range("A2:CZ" & coll).Select
Selection.ClearContents
Range("A1").Select
Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=True
Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\PLANTILLA_HC")
xiaom = "PLANTILLA_HC"
st = xiaom & ".xlsx"
Sheets(1).Select
Range("A1").Select
cooll = Selection.End(xlDown).Row
Sheets(1).Range("A2:U" & cooll).Select
Selection.ClearContents
Range("A1").Select
Windows(st).Activate
ActiveWorkbook.Close SaveChanges:=True
'Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\NPS\NPS")
'xiom = "NPS_CED_PLANTILLA"
'tt = xiom & ".xlsx"
'Sheets(1).Select
'Range("A1").Select
'ActiveSheet.ShowAllData
'Range("A1").Select
'coooll = Selection.End(xlDown).Row
'Sheets(1).Range("A1:E" & coooll).Select
'Selection.ClearContents
'Range("A1").Select
'Windows(tt).Activate
'ActiveWorkbook.Close SaveChanges:=True
Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\VIEJA\HC_VIEJO")
eg = "HC_VIEJO"
xx = eg & ".xlsx"
Sheets("HC").Select
Range("A1").Select
ActiveSheet.ShowAllData
Range("A1").Select
ty = Selection.End(xlDown).Row
Sheets("HC").Range("A2:AB" & ty).Select
Selection.ClearContents
Range("AF1").ClearContents
Range("AI1").ClearContents
Range("A1").Select
Sheets("NUEVA").Delete
Windows(xx).Activate
ActiveWorkbook.Close SaveChanges:=True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub pinot()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos

Dim vari_2 As Excel.Workbook

Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\NPS\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\NPS\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate

Sheets("CALCULO").Select

Range("H3").Select
ty = Selection.End(xlDown).Row - 1
Range("H3:H" & ty).Select
Selection.Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("NPS-UMBRAL").Select
Range("AH2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows(tt).Activate
Range("L3:L" & ty).Select
Selection.Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("NPS-UMBRAL").Select
Range("AI2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
Range("AI2").Select

MsgBox "LISTO EL PRIMER PEGUE", vbInformation


MsgBox "Fase inicial completada", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True



End Sub

Sub pinarello()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos

Dim vari_2 As Excel.Workbook

Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Dim top1, top2, assd As String

archivos = Dir("D:\AUTOMATIZACION\UMBRAL\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\UMBRAL\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1



Windows(tt).Activate
Sheets("Data").Select
variab_1 = Sheets("Data").Range("A" & Rows.Count).End(xlUp).Row
   Sheets("Data").Range("A2:A" & variab_1).Copy

Windows("PLAN_LIQ.xlsm").Activate
Sheets("NPS-UMBRAL").Select
Range("A2").Select
ActiveSheet.Paste
Application.CutCopyMode = False

'++
Windows(tt).Activate
Sheets("Data").Select
Sheets("Data").Range("E2:E" & variab_1).Copy


Windows("PLAN_LIQ.xlsm").Activate
Sheets("NPS-UMBRAL").Select
Range("E2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
'++++

Windows(tt).Activate
Sheets("Data").Select
Sheets("Data").Range("J2:J" & variab_1).Copy


Windows("PLAN_LIQ.xlsm").Activate
Sheets("NPS-UMBRAL").Select
Range("J2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
'++++



Windows(tt).Activate
Sheets("Data").Select
Sheets("Data").Range("L2:L" & variab_1).Copy


Windows("PLAN_LIQ.xlsm").Activate
Sheets("NPS-UMBRAL").Select
Range("L2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "Listo el pegue", vbInformation
'++++
asd = Dir("D:\AUTOMATIZACION\NPS\*.xlsx")
top1 = "D:\AUTOMATIZACION\NPS\" & asd: top2 = "D:\AUTOMATIZACION\NPS\NPS.xlsx"
Name top1 As top2
MsgBox "LISTO PRIMERA FASE", vbInformation
Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\NPS\NPS")
ills = ActiveWorkbook.Name
gg = ills
Windows("PLAN_LIQ.xlsm").Activate
Range("A1").Select
ty = Selection.End(xlDown).Row - 1
Range("O2").Select
For i = 1 To ty
ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-14],[NPS.xlsx]CALCULO!C8:C13,6,FALSE),""NO REGISTRA"")"
ActiveCell.Offset(1, 0).Select
Next i
Windows(gg).Activate
ActiveWorkbook.Close SaveChanges:=False
Windows("PLAN_LIQ.xlsm").Activate
Columns("O:O").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("O1").Select
Application.CutCopyMode = False

MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True


End Sub
Sub gatti()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos

Dim vari_2 As Excel.Workbook

Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\UMBRAL\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\UMBRAL\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate
Sheets("Data").Select
variab_1 = Sheets("Data").Range("A" & Rows.Count).End(xlUp).Row
   Sheets("Data").Range("O2:P" & variab_1).Copy

Windows("PLAN_LIQ.xlsm").Activate
Sheets("NPS-UMBRAL").Select
Range("E2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub tourmalet()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Range("AQ3").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Selection.Font.Bold = False
     Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-33]<=0,0,IFERROR(VLOOKUP(RC[-40],'NPS-UMBRAL'!C42:C45,4,0)*(RC[-33]/RC[28]),""SIN OFC""))"
    Range("AQ3").Select
   Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    ActiveSheet.Paste
    Range("AQ3").Select
    Application.CutCopyMode = False
    Range("AQ3") = "PPTO UMBRAL PROPOR"
    Range("AQ3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
   Range("A1").Select
ActiveSheet.ShowAllData
Columns("BM:BM").Select
Selection.ClearContents
Range("A1").Select
MsgBox "listo segundo cálculo", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub cosa_nostra()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Select
vari_9 = Range("a1").Value
Range("BT4").Select
For i = 1 To vari_9
ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-69],INCENTIVO!C1:C4,4,FALSE),0)"
ActiveCell.Offset(1, 0).Select
Next i
Range("A1").Select
MsgBox "listo primer cálculo", vbInformation
Sheets("Liquidacion").Select
zidan = Range("a1").Value
Range("CE4").Select
For i = 1 To zidan
ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-80],INCENTIVO!C1:C4,4,FALSE),0)"
ActiveCell.Offset(1, 0).Select
Next i
Range("A1").Select
MsgBox "listo SEGUNDO cálculo", vbInformation
Sheets("ASESORES TMK").Select
idan = Range("a1").Value
Range("AJ4").Select
For i = 1 To idan
 ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-33],INCENTIVO!C1:C4,4,FALSE),0)"
ActiveCell.Offset(1, 0).Select
Next i
Range("A1").Select
MsgBox "listo TERCER cálculo", vbInformation
Sheets("ROL-APP").Select
dan = Range("a1").Value
Range("BK4").Select
For i = 1 To dan
ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-54],INCENTIVO!C1:C4,4,FALSE),0)"
ActiveCell.Offset(1, 0).Select
Next i
Range("A1").Select
MsgBox "listo CUARTO cálculo", vbInformation
Sheets("OTRAS GESTIONES").Select
dann = Range("a1").Value
Range("AC4").Select
For i = 1 To dann
ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-26],INCENTIVO!C1:C4,4,FALSE),0)"
ActiveCell.Offset(1, 0).Select
Next i
Range("A1").Select
MsgBox "listo QUINTO cálculo", vbInformation
'++
Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Select
vari_9 = Range("a1").Value
Range("AZ4").Select
For i = 1 To vari_9
ActiveCell.FormulaR1C1 = "=SUM(RC[-3],RC[-18],RC[-1],RC[18],RC[20])"
ActiveCell.Offset(1, 0).Select
Next i
Range("A1").Select
MsgBox "listo primer cálculo_V2", vbInformation
Sheets("Liquidacion").Select
zidan = Range("a1").Value
Range("BK4").Select
For i = 1 To zidan
    ActiveCell.FormulaR1C1 = _
        "=SUM((IF((AND(RC[-26]=0,OR(RC[-19]=0,RC[-19]=""""))),0,SUM(RC[-3],RC[-29],RC[-1]))),RC[18],RC[20])"
ActiveCell.Offset(1, 0).Select
Next i
Range("A1").Select
MsgBox "listo SEGUNDO cálculo_V2", vbInformation
Sheets("ASESORES TMK").Select
idan = Range("a1").Value
Range("R4").Select
For i = 1 To idan
    ActiveCell.FormulaR1C1 = _
        "=SUM(((RC[-2]*RC[-10])*(RC[-8]/RC[6])),RC[16],RC[18])"
ActiveCell.Offset(1, 0).Select
Next i
Range("A1").Select
MsgBox "listo TERCER cálculo_V2", vbInformation
Sheets("ROL-APP").Select
dan = Range("a1").Value
Range("AT4").Select
For i = 1 To dan
    ActiveCell.FormulaR1C1 = _
        "=SUM(((RC[-1]*RC[-30])*(RC[-28]/RC[7])),RC[17],RC[19])"
ActiveCell.Offset(1, 0).Select
Next i
Range("A1").Select
MsgBox "listo CUARTO cálculo_V2", vbInformation
Sheets("OTRAS GESTIONES").Select
dann = Range("a1").Value
Range("K4").Select
For i = 1 To dann
    ActiveCell.FormulaR1C1 = "=((RC[-5]*80%)*(RC[-3]/RC[6]))+RC[-2]+RC[16]+RC[18]"
ActiveCell.Offset(1, 0).Select
Next i
Range("A1").Select
MsgBox "listo QUINTO cálculo_V2", vbInformation

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation

End Sub
Sub mutar()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim asd As String
Dim top1, top2, assd As String
asd = Dir("D:\AUTOMATIZACION\NPS\*.csv")
top1 = "D:\AUTOMATIZACION\NPS\" & asd: top2 = "D:\AUTOMATIZACION\NPS\NPS_CEDULA.csv"
Name top1 As top2
MsgBox "LISTO PRIMERA FASE", vbInformation
assd = Dir("D:\AUTOMATIZACION\NPS\USUARIO\*.csv")
top1 = "D:\AUTOMATIZACION\NPS\USUARIO\" & assd: top2 = "D:\AUTOMATIZACION\NPS\USUARIO\NPS_USUARIORED.csv"
Name top1 As top2
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub ejea()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\NPS\NPS")
vari_1 = ActiveWorkbook.Name
tt = vari_1
Windows(tt).Activate
Sheets("CALCULO").Select
Range("N2") = "CRUCE"
Range("H2").Select
va_8 = Selection.End(xlDown).Row - 1
Range("N3").Select
ActiveCell.FormulaR1C1 = _
"=VLOOKUP(RC[-6],'[PLAN_LIQ.xlsm]NPS-UMBRAL'!C1,1,FALSE)"
ActiveCell.Copy
Range("N3:N" & va_8).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("N2").Select
MsgBox "Listo primera fase", vbInformation
Range("H2:N2").Select
Selection.AutoFilter
ActiveSheet.ListObjects.Add(xlSrcRange, Range("H2:N" & va_8), , xlYes).Name = "Tabla1"
Set variable_1 = Sheets("CALCULO").Range("N:N").Find(What:="#N/A", LookIn:=xlValues, LookAt:=xlWhole)
If Not variable_1 Is Nothing Then
Application.Goto Reference:="Tabla1"
Selection.AutoFilter Field:=7, Criteria1:="#N/A"
Range("H2").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("NPS-UMBRAL").Select
Range("A1").Select
zlata = Selection.End(xlDown).Row + 1
Range("A" & zlata).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Windows(tt).Activate
Range("M2").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("NPS-UMBRAL").Select
Range("O" & zlata).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
End If
If variable_1 Is Nothing Then
MsgBox "TODOS LOS DATOS CRUZARON"
End If
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub marado()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
archivos = Dir("D:\AUTOMATIZACION\4_MES\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\4_MES\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate
Sheets("TABLAS").Select
col = Sheets("TABLAS").Range("A" & Rows.Count).End(xlUp).Row
Range("A3:E" & col).Select
Selection.Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("MOVIL").Select
Range("BA3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "LISTO PARTE 1", vbInformation
Range("AJ1").Select
ActiveCell.FormulaR1C1 = _
        "=UPPER((CONCATENATE(""PLAN POWER 4 MES - EVALUACIÓN VENTAS MES DE: "",TEXT(TODAY()-120,""MMMM""))))"
MsgBox "LISTO PRIMERA FASE", vbInformation


archivos = Dir("D:\AUTOMATIZACION\CLARO PAY\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\CLARO PAY\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate

Sheets(1).Select

col = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Range("A1:B" & col).Select
Selection.Copy

Windows("PLAN_LIQ.xlsm").Activate
Sheets("Fuente Hogares").Select
Range("BV5").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub python()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos

Dim vari_2 As Excel.Workbook

Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\METAS\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\METAS\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate

Sheets(1).Select
vegas = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
Sheets(1).Range("B2:B" & vegas).Copy
    
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("AO3").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Windows(tt).Activate
Sheets(1).Range("R2:R" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("AP3").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Windows(tt).Activate
Sheets(1).Range("T2:T" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("AQ3").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Windows(tt).Activate
Sheets(1).Range("Y2:Z" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("AR3").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Windows(tt).Activate
Sheets(1).Range("X2:X" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("AT3").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Windows(tt).Activate
Sheets(1).Range("AA2:AA" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("AU3").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Windows(tt).Activate
Sheets(1).Range("AD2:AD" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Metas").Select
Range("AV3").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub
Sub stromae()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Sheets("HC").Select
Range("C1").Select
vari_7 = Selection.End(xlDown).Row
Set rangodatos = Sheets("HC").Range("A1:AB" & vari_7)
rangodatos.AutoFilter Field:=28, Criteria1:="*"
rangodatos.AutoFilter Field:=27, Criteria1:="="
rangodatos.AutoFilter Field:=26, Criteria1:="=Cav App", Operator:=xlOr, Criteria2:="=Cav Tp"
rangodatos.AutoFilter Field:=25, Criteria1:="#N/A"
Range("C1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("ROL-APP").Select
Range("C4").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("C4").Select
MsgBox "listo cédulas_1", vbInformation
Sheets("HC").Select
Range("C1").Select
ActiveSheet.ShowAllData
rangodatos.AutoFilter Field:=28, Criteria1:="*"
rangodatos.AutoFilter Field:=27, Criteria1:="="
rangodatos.AutoFilter Field:=26, Criteria1:="Cav Barra"
rangodatos.AutoFilter Field:=25, Criteria1:="#N/A"
rangodatos.AutoFilter Field:=2, Criteria1:="#N/A"
Range("C1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("Liquidacion").Select
Range("C4").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("C4").Select
MsgBox "listo cédulas_2", vbInformation
'Sheets("HC").Select
'Range("C1").Select
'ActiveSheet.ShowAllData
'rangodatos.AutoFilter Field:=28, Criteria1:="*"
'rangodatos.AutoFilter Field:=27, Criteria1:="="
'rangodatos.AutoFilter Field:=26, Criteria1:="Cav Barra"
'rangodatos.AutoFilter Field:=25, Criteria1:="#N/A"
'rangodatos.AutoFilter Field:=2, Criteria1:="<>#N/A"
'Range("C1").Select
'ActiveCell.Offset(1, 0).Select
'Range(Selection, Selection.End(xlDown)).Select
'Selection.Copy
'Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Select
'Range("C4").Select
'ActiveSheet.Paste
'Application.CutCopyMode = False
'Range("C4").Select
MsgBox "listo cédulas_3", vbInformation
Sheets("HC").Select
Range("C1").Select
ActiveSheet.ShowAllData
rangodatos.AutoFilter Field:=28, Criteria1:="*"
rangodatos.AutoFilter Field:=27, Criteria1:="="
rangodatos.AutoFilter Field:=26, Criteria1:="Cav Ev"
Range("C1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Select
'Range("C3").Select
'vari_7 = Selection.End(xlDown).Row
'Range("C" & vari_7 + 1).Select
Range("C4").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("C4").Select
MsgBox "listo cédulas_4", vbInformation
Sheets("HC").Select
Range("C1").Select
ActiveSheet.ShowAllData
rangodatos.AutoFilter Field:=28, Criteria1:="*"
rangodatos.AutoFilter Field:=27, Criteria1:="="
rangodatos.AutoFilter Field:=26, Criteria1:="Otras Gestiones"
Range("C1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("OTRAS GESTIONES").Select
Range("C4").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("C4").Select
MsgBox "listo cédulas_5", vbInformation
Sheets("HC").Select
Range("C1").Select
ActiveSheet.ShowAllData
rangodatos.AutoFilter Field:=28, Criteria1:="*"
rangodatos.AutoFilter Field:=27, Criteria1:="="
rangodatos.AutoFilter Field:=26, Criteria1:="Tmk"
Range("C1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("ASESORES TMK").Select
Range("C4").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("C4").Select
MsgBox "listo cédulas_6", vbInformation
Sheets("HC").Select
Range("C1").Select
ActiveSheet.ShowAllData
rangodatos.AutoFilter Field:=28, Criteria1:="*"
rangodatos.AutoFilter Field:=27, Criteria1:="="
rangodatos.AutoFilter Field:=26, Criteria1:="Garantizado Inspira Gh"
rangodatos.AutoFilter Field:=7, Criteria1:=Array("Consultor Integral Servicio A Clientes", "Asesor Servicio Al Cliente", "Consultor Integral Servicio A Clientes Sr", "Asesor Integral Servicio Al Cliente", "Consultor(a) Integral Servicio A Clientes", "Asesor(a) Servicio Al Cliente", "Consultor(a) Integral Servicio A Clientes Sr", "Consultor(a) Servicio Personalizado A Clientes", "Asesor(a) Integral Servicio Al Cliente"), Operator:=xlFilterValues
Range("C1").Select
ActiveCell.Offset(1, 0).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("Consultores_Inspira").Select
Range("C4").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("C4").Select
MsgBox "listo cédulas_7", vbInformation
Sheets("HC").Select
Range("C1").Select
ActiveSheet.ShowAllData
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub polar()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Sheets("Liquidacion").Select
Range("C4").Select
Range(Selection, Selection.End(xlDown)).Copy
Sheets("Resumen Puntos").Select
Range("A6").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Sheets("Liquid Tiendas-Cvc-DENTRO CAV").Select
Range("C4").Select
Range(Selection, Selection.End(xlDown)).Copy
Sheets("Resumen Puntos").Select
Range("A6").Select
vari_7 = Selection.End(xlDown).Row
Range("A" & vari_7 + 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A6").Select
MsgBox "Listo primera fase", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub roubaix()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos

Dim vari_2 As Excel.Workbook

Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String

archivos = Dir("D:\AUTOMATIZACION\EFICIENCIA\*.xlsx")

Do While archivos <> ""
Workbooks.Open "D:\AUTOMATIZACION\EFICIENCIA\" & archivos

archivos = Dir
Loop

vari_1 = ActiveWorkbook.Name

tt = vari_1

Windows(tt).Activate

Sheets("Consultor").Select
vegas = Sheets("Consultor").Range("B" & Rows.Count).End(xlUp).Row
Sheets("Consultor").Range("B5:B" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("NPS-UMBRAL").Select
Range("BC2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Windows(tt).Activate
Sheets("Consultor").Range("E5:E" & vegas).Copy
Windows("PLAN_LIQ.xlsm").Activate
Sheets("NPS-UMBRAL").Select
Range("BD2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Windows(tt).Activate
ActiveWorkbook.Close SaveChanges:=False
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
    
End Sub
Sub veratti()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

On Error Resume Next

Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String


Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\FIJAS SOLO CAV")


xiaomi = "FIJAS SOLO CAV"
    ss = xiaomi & ".xlsx"

Windows(ss).Activate


Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)

Worksheets.Add
ActiveSheet.Name = "Hoja de pivote4"

Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("A1"))


With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").AddDataField ActiveSheet. _
        PivotTables("Tabla dinámica1").PivotFields("ValReg"), "Cuenta de ValReg", _
        xlCount
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Renta actual")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Producto Liquid")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("ValArpu")
        .Orientation = xlPageField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Back_Lock")
        .Orientation = xlPageField
        .Position = 4
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Back_Lock"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Back_Lock"). _
        CurrentPage = "(blank)"
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Renta actual"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Renta actual")
        .PivotItems("0").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Renta actual"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("ValArpu"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("ValArpu").CurrentPage _
        = "S"
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Producto Liquid"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Producto Liquid")
        .PivotItems("@DTH").Visible = False
        .PivotItems("100M").Visible = False
        .PivotItems("100MFO").Visible = False
        .PivotItems("120M").Visible = False
        .PivotItems("140M").Visible = False
        .PivotItems("150M").Visible = False
        .PivotItems("150MFO").Visible = False
        .PivotItems("15M").Visible = False
        .PivotItems("160M").Visible = False
        .PivotItems("200M").Visible = False
        .PivotItems("20M").Visible = False
        .PivotItems("210M").Visible = False
        .PivotItems("23M").Visible = False
        .PivotItems("240M").Visible = False
        .PivotItems("300M").Visible = False
        .PivotItems("30M").Visible = False
        .PivotItems("310M").Visible = False
        .PivotItems("400M").Visible = False
        .PivotItems("40M").Visible = False
        .PivotItems("45M").Visible = False
        .PivotItems("50M").Visible = False
        .PivotItems("60M").Visible = False
        .PivotItems("75M").Visible = False
        .PivotItems("80M").Visible = False
        .PivotItems("8M").Visible = False
        .PivotItems("CV").Visible = False
        .PivotItems("DTH").Visible = False
        .PivotItems("DTHA").Visible = False
        .PivotItems("DTHS").Visible = False
        .PivotItems("FPFOX").Visible = False
        .PivotItems("GOLD").Visible = False
        .PivotItems("HBO").Visible = False
        .PivotItems("HBOMX").Visible = False
        .PivotItems("HD").Visible = False
        .PivotItems("HPC").Visible = False
        .PivotItems("IO").Visible = False
        .PivotItems("NVA").Visible = False
        .PivotItems("PC").Visible = False
        .PivotItems("PCI").Visible = False
        .PivotItems("R15").Visible = False
        .PivotItems("RJ").Visible = False
        .PivotItems("SP").Visible = False
        .PivotItems("TDB").Visible = False
        .PivotItems("TDP").Visible = False
        .PivotItems("TDS").Visible = False
        .PivotItems("TEL").Visible = False
        .PivotItems("TELDTH").Visible = False
        .PivotItems("TV").Visible = False
        .PivotItems("UW75L").Visible = False
        .PivotItems("UWARR").Visible = False
        .PivotItems("VC").Visible = False
    End With
    With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Producto Liquid")
        .PivotItems("VEL").Visible = False
        .PivotItems("VEL2").Visible = False
        .PivotItems("VEL3").Visible = False
        .PivotItems("VEL4").Visible = False
        .PivotItems("WIFI").Visible = False
        .PivotItems("WINP").Visible = False
        .PivotItems("WINPSD").Visible = False
        .PivotItems("WINSD").Visible = False
        .PivotItems("(blank)").Visible = False
        .PivotItems("NFXSD").Visible = False
        .PivotItems("NFXPRE").Visible = False
        
    End With
    ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("Producto Liquid"). _
        EnableMultiplePageItems = True

MsgBox "listo basico", vbInformation
'++
Windows(ss).Activate

Sheets("Hoja de pivote4").Select

Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)


Sheets("Hoja de pivote4").Select
Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("D1"))



With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").AddDataField ActiveSheet. _
        PivotTables("Tabla dinámica2").PivotFields("ValReg"), "Cuenta de ValReg", _
        xlCount
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Renta actual")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Producto Liquid")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("ValArpu")
        .Orientation = xlPageField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Back_Lock")
        .Orientation = xlPageField
        .Position = 4
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Back_Lock"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Back_Lock"). _
        CurrentPage = "(blank)"
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Renta actual"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Renta actual")
        .PivotItems("0").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Renta actual"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("ValArpu"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("ValArpu").CurrentPage _
        = "S"
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Producto Liquid"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Producto Liquid")
        .PivotItems("@DTH").Visible = False
        .PivotItems("100M").Visible = False
        .PivotItems("100MFO").Visible = False
        .PivotItems("120M").Visible = False
        .PivotItems("140M").Visible = False
        .PivotItems("150M").Visible = False
        .PivotItems("150MFO").Visible = False
        .PivotItems("15M").Visible = False
        .PivotItems("160M").Visible = False
        .PivotItems("200M").Visible = False
        .PivotItems("20M").Visible = False
        .PivotItems("210M").Visible = False
        .PivotItems("23M").Visible = False
        .PivotItems("240M").Visible = False
        .PivotItems("300M").Visible = False
        .PivotItems("30M").Visible = False
        .PivotItems("310M").Visible = False
        .PivotItems("400M").Visible = False
        .PivotItems("40M").Visible = False
        .PivotItems("45M").Visible = False
        .PivotItems("50M").Visible = False
        .PivotItems("60M").Visible = False
        .PivotItems("75M").Visible = False
        .PivotItems("80M").Visible = False
        .PivotItems("8M").Visible = False
        .PivotItems("CV").Visible = False
        .PivotItems("DTH").Visible = False
        .PivotItems("DTHA").Visible = False
        .PivotItems("DTHS").Visible = False
        .PivotItems("FPFOX").Visible = False
        .PivotItems("GOLD").Visible = False
        .PivotItems("HBO").Visible = False
        .PivotItems("HBOMX").Visible = False
        .PivotItems("HD").Visible = False
        .PivotItems("HPC").Visible = False
        .PivotItems("IO").Visible = False
        .PivotItems("NVA").Visible = False
        .PivotItems("PC").Visible = False
        .PivotItems("PCI").Visible = False
        .PivotItems("R15").Visible = False
        .PivotItems("RJ").Visible = False
        .PivotItems("SP").Visible = False
        .PivotItems("TDB").Visible = False
        .PivotItems("TDP").Visible = False
        .PivotItems("TDS").Visible = False
        .PivotItems("TEL").Visible = False
        .PivotItems("TELDTH").Visible = False
        .PivotItems("TV").Visible = False
        .PivotItems("UW75L").Visible = False
        .PivotItems("UWARR").Visible = False
        .PivotItems("VC").Visible = False
        
    End With
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Producto Liquid")
        .PivotItems("VEL").Visible = False
        .PivotItems("VEL2").Visible = False
        .PivotItems("VEL3").Visible = False
        .PivotItems("VEL4").Visible = False
        .PivotItems("WIFI").Visible = False
        .PivotItems("WINP").Visible = False
        .PivotItems("WINPSD").Visible = False
        .PivotItems("WINSD").Visible = False
        .PivotItems("(blank)").Visible = False
        .PivotItems("TRHP").Visible = False
        .PivotItems("TRHB").Visible = False
        .PivotItems("TRHO").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("Producto Liquid"). _
        EnableMultiplePageItems = True
    
'++
Windows(ss).Activate

Sheets("Hoja de pivote4").Select

Set pc = ActiveWorkbook.PivotCaches.Create(xlDatabase, Sheets("Hoja1").Range("A1").CurrentRegion.Address)


Sheets("Hoja de pivote4").Select
Set pt = ActiveSheet.PivotTables.Add(PivotCache:=pc, TableDestination:=Range("H1"))



With ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Cedula asesor")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabla dinámica3").AddDataField ActiveSheet. _
        PivotTables("Tabla dinámica3").PivotFields("ValReg"), "Cuenta de ValReg", _
        xlCount
    With ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Renta actual")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Producto Liquid")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("ValArpu")
        .Orientation = xlPageField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Back_Lock")
        .Orientation = xlPageField
        .Position = 4
    End With
    ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Back_Lock"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Back_Lock"). _
        CurrentPage = "(blank)"
    ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Renta actual"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Renta actual")
        .PivotItems("0").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Renta actual"). _
        EnableMultiplePageItems = True
    ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("ValArpu"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("ValArpu").CurrentPage _
        = "S"
    ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Producto Liquid"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Producto Liquid")
        .PivotItems("@DTH").Visible = False
        .PivotItems("100M").Visible = False
        .PivotItems("100MFO").Visible = False
        .PivotItems("120M").Visible = False
        .PivotItems("140M").Visible = False
        .PivotItems("150M").Visible = False
        .PivotItems("150MFO").Visible = False
        .PivotItems("15M").Visible = False
        .PivotItems("160M").Visible = False
        .PivotItems("200M").Visible = False
        .PivotItems("20M").Visible = False
        .PivotItems("210M").Visible = False
        .PivotItems("23M").Visible = False
        .PivotItems("240M").Visible = False
        .PivotItems("300M").Visible = False
        .PivotItems("30M").Visible = False
        .PivotItems("310M").Visible = False
        .PivotItems("400M").Visible = False
        .PivotItems("40M").Visible = False
        .PivotItems("45M").Visible = False
        .PivotItems("50M").Visible = False
        .PivotItems("60M").Visible = False
        .PivotItems("75M").Visible = False
        .PivotItems("80M").Visible = False
        .PivotItems("8M").Visible = False
        .PivotItems("CV").Visible = False
        .PivotItems("DTH").Visible = False
        .PivotItems("DTHA").Visible = False
        .PivotItems("DTHS").Visible = False
        .PivotItems("FPFOX").Visible = False
        .PivotItems("GOLD").Visible = False
        .PivotItems("HBO").Visible = False
        .PivotItems("HBOMX").Visible = False
        .PivotItems("HD").Visible = False
        .PivotItems("HPC").Visible = False
        .PivotItems("IO").Visible = False
        .PivotItems("NVA").Visible = False
        .PivotItems("PC").Visible = False
        .PivotItems("PCI").Visible = False
        .PivotItems("R15").Visible = False
        .PivotItems("RJ").Visible = False
        .PivotItems("SP").Visible = False
        .PivotItems("TDB").Visible = False
        .PivotItems("TDP").Visible = False
        .PivotItems("TDS").Visible = False
        .PivotItems("TEL").Visible = False
        .PivotItems("TELDTH").Visible = False
        .PivotItems("TV").Visible = False
        .PivotItems("UW75L").Visible = False
        .PivotItems("UWARR").Visible = False
        .PivotItems("VC").Visible = False
        
    End With
    With ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Producto Liquid")
        .PivotItems("VEL").Visible = False
        .PivotItems("VEL2").Visible = False
        .PivotItems("VEL3").Visible = False
        .PivotItems("VEL4").Visible = False
        .PivotItems("WIFI").Visible = False
        .PivotItems("WINP").Visible = False
        .PivotItems("WINPSD").Visible = False
        .PivotItems("WINSD").Visible = False
        .PivotItems("(blank)").Visible = False
        .PivotItems("NFXSD").Visible = False
        .PivotItems("NFXPRE").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica3").PivotFields("Producto Liquid"). _
        EnableMultiplePageItems = True
    

'++


Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=True
Call gallardo
Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA ABRIR EL ARCHIVO Y EJECUTE EL BOTON MAGIA", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.Quit


End Sub


Sub gallardo()
Application.ScreenUpdating = False
Application.DisplayAlerts = False



Dim vari_2 As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Dim variable_1 As Object




Set vari_2 = Workbooks.Open("D:\AUTOMATIZACION\PLANTILLAS\FIJAS SOLO CAV")


xiaomi = "FIJAS SOLO CAV"
    ss = xiaomi & ".xlsx"



Windows(ss).Activate

Sheets("Hoja1").Select
Range("A1").Select

Set variable_1 = Sheets("Hoja1").Range("AK:AK").Find(What:="NFX", LookIn:=xlValues, LookAt:=xlWhole)

If Not variable_1 Is Nothing Then
Sheets("Hoja de pivote4").Select
ActiveSheet.PivotTables("Tabla dinámica1").PivotSelect "", xlDataAndLabel, True
Selection.Copy


Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("BS4").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Windows(ss).Activate
Sheets("Hoja de pivote4").Select
ActiveSheet.PivotTables("Tabla dinámica2").PivotSelect "", xlDataAndLabel, True
Selection.Copy


Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("BO4").Select
ActiveSheet.Paste
Application.CutCopyMode = False


'+
Windows(ss).Activate
Sheets("Hoja de pivote4").Select
ActiveSheet.PivotTables("Tabla dinámica3").PivotSelect "", xlDataAndLabel, True
Selection.Copy


Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("BL4").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A1").Select



MsgBox "listo NFX_tabla", vbInformation
End If
 If variable_1 Is Nothing Then
 
 Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("BS4") = "NO SE REGISTRAN VENTAS NFX"

Windows(ss).Activate
Sheets("Hoja de pivote4").Select
ActiveSheet.PivotTables("Tabla dinámica2").PivotSelect "", xlDataAndLabel, True
Selection.Copy


Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("BO4").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A1").Select

'+
Windows(ss).Activate
Sheets("Hoja de pivote4").Select
ActiveSheet.PivotTables("Tabla dinámica3").PivotSelect "", xlDataAndLabel, True
Selection.Copy


Windows("PLAN_LIQ.xlsm").Activate

Sheets("Fuente Hogares").Select
Range("BL4").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A1").Select

MsgBox "listo NFX_no", vbInformation
End If
Windows(ss).Activate
ActiveWorkbook.Close SaveChanges:=False

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub
Sub gorosito()


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Fuente Hogares").Select
Columns("P:Q").Select
Selection.Copy
Columns("P:Q").Select
Application.CutCopyMode = False
Selection.Copy
Columns("S:S").Select
ActiveSheet.Paste
Columns("V:V").Select
ActiveSheet.Paste
Range("S5").Select
Range("S5") = "VELOCIDAD 3"
Range("V5").Select
Range("V5") = "VELOCIDAD 4"
Application.CutCopyMode = False
Range("P1").Select
MsgBox "LISTO PRIMERA FASE", vbInformation
Range("T6").Select
ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("Producto Liquid").ClearAllFilters
ActiveSheet.PivotTables("Tabla dinámica12").PivotFields("Producto Liquid").CurrentPage = "VEL3"
Range("W6").Select
ActiveSheet.PivotTables("Tabla dinámica13").PivotFields("Producto Liquid").ClearAllFilters
ActiveSheet.PivotTables("Tabla dinámica13").PivotFields("Producto Liquid").CurrentPage = "VEL4"
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA ABRIR EL ARCHIVO Y EJECUTE EL BOTON MAGIA", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.Quit
End Sub
Sub nft()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("Tabla Var").Select
va_7 = Range("U1").Value
Sheets("Consultores_Inspira").Select
Range("G2") = va_7
Range("a1").Select
ActiveCell.FormulaR1C1 = "=COUNTA(C[2])-2"
vari_9 = Range("a1").Value
Range("c4").Select

For i = 1 To vari_9
 ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],HC!C3:C5,3,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],HC!C3:C7,5,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],HC!C3:C9,7,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R2C7,Meses!C1:C2,2,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-5],'DESARROLLO+PROYECTOS'!C3:C21,19,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[3]>0,RC[-1]>0,RC[-1]<RC[3]),0,IF(AND(RC[3]>0,RC[-1]>0,RC[-1]>RC[3]),RC[-1]-RC[3],RC[-1]))"
    ActiveCell.Offset(0, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-9],Garantizado!C3:C5,3,0),0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-10],HC!C3:C20,18,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-11],HC!C3:C12,10,0)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-12],HC!C3:C26,24,FALSE)"
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlToLeft).Select
    
Next i

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation

End Sub


Sub FORMATITOS_21()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Sheets("Consultores_Inspira").Select
Range("a1").Select
vari_9 = Range("a1").Value
Range("A4").Select
For i = 1 To vari_9

 ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[2],HC!C3:C13,11,FALSE)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],HC!C3:C4,2,0)"
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlToLeft).Select
Next i

MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
End Sub

Sub valdano()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Fuente Hogares").Select
Columns("A:B").Select
Selection.Copy
Range("BZ1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("BZ1").Select
ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("NO TIENE APP = ""X""") _
    .CurrentPage = "(All)"
With ActiveSheet.PivotTables("Tabla dinámica1").PivotFields( _
    "NO TIENE APP = ""X""")
    .PivotItems("X").Visible = False
End With
ActiveSheet.PivotTables("Tabla dinámica1").PivotFields("NO TIENE APP = ""X""") _
    .EnableMultiplePageItems = True
MsgBox "listo primera fase", vbInformation
Columns("BZ:CA").Select
Selection.Copy
Range("CD1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("CD1").Select
ActiveSheet.PivotTables("Tabla dinámica20").PivotFields("Tipo producto"). _
    CurrentPage = "(All)"
With ActiveSheet.PivotTables("Tabla dinámica20").PivotFields("Tipo producto")
    .PivotItems("R").Visible = False
End With
ActiveSheet.PivotTables("Tabla dinámica20").PivotFields("Tipo producto"). _
    EnableMultiplePageItems = True
ActiveSheet.PivotTables("Tabla dinámica20").PivotFields("NO TIENE APP = ""X""") _
    .Orientation = xlHidden
Range("BZ2") = "ALTAS MASIVO HOGAR CON APP"
Range("CD2") = "ALTAS NEGOCIOS HOGAR TOTAL"
Range("A2") = "ALTAS MASIVO HOGAR SIN APP"
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA ABRIR EL ARCHIVO Y EJECUTE EL BOTON MAGIA", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.Quit
End Sub

Sub chicanos()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Fuente Hogares").Select
Columns("L:M").Select
Selection.Copy
Range("BC1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("BC2") = "CANALES DEPORTIVOS"
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA ABRIR EL ARCHIVO Y EJECUTE EL BOTON MAGIA", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.Quit
End Sub
Sub klopo()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLAN_LIQ.xlsm").Activate
Sheets("Fuente Hogares").Select
Columns("CD:CE").Select
Selection.Copy
Range("CI1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("CI1").Select
Range("CI2") = "EJECUCIÓN INTERNET"
ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("Tipo producto"). _
Orientation = xlHidden
With ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("Producto Venta")
.Orientation = xlPageField
.Position = 4
End With
ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("Producto Venta"). _
CurrentPage = "(All)"
ActiveSheet.PivotTables("Tabla dinámica10").PivotFields("Producto Venta"). _
EnableMultiplePageItems = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
MsgBox "EL SISTEMA SE CERRARÁ, POR FAVOR NO MANIPULE EL COMPUTADOR HASTA QUE EL EJECUTABLE DE EXCEL SE CIERRE, VUELVA ABRIR EL ARCHIVO Y EJECUTE EL BOTON MAGIA", vbInformation
ActiveWorkbook.Close SaveChanges:=True
Application.Quit
End Sub

