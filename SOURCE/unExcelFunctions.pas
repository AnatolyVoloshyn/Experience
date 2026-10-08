unit unExcelFunctions;

interface
uses dxSpreadSheet, dxSpreadSheetCore, System.Types, System.Math, dxSpreadSheetGraphics, System.UITypes,
     cxGraphics, dxSpreadSheetCoreHistory, dxSpreadSheetCoreStyles, dxSpreadSheetCoreStrs,
     dxSpreadSheetConditionalFormatting, dxSpreadSheetConditionalFormattingRules,
     dxSpreadSheetClasses, dxSpreadSheetContainers, dxSpreadSheetFormulas,
     dxSpreadSheetHyperlinks, dxSpreadSheetFunctions, dxSpreadSheetStyles,
     dxSpreadSheetPrinting, dxSpreadSheetTypes,
     dxSpreadSheetUtils, dxSpreadSheetFormattedTextUtils, System.SysUtils, dxHashUtils;


 var arr_hiperlinks:array of array of variant;
 var temp_SS:TdxSpreadSheet;

 function  AddWorksheet(SS:TdxSpreadSheet; new_caption:string):integer;
 procedure RenameWorksheet(SS:TdxSpreadSheet; Worksheet:integer; new_caption:string);
 procedure MergeCells(SS:TdxSpreadSheet; Worksheet,nLeft,nTop,nRight,nBottom:integer);
 procedure SetColumnWidth(SS:TdxSpreadSheet; Worksheet:integer; Column:integer; Width:Integer);
 procedure SetRowHaight(SS:TdxSpreadSheet; Worksheet:integer; Row:integer; Haight:Integer);

 procedure SetColumnAutoBestWidth(SS:TdxSpreadSheet; Worksheet:integer; Column:integer);
 procedure SetRowAutoBestHaight(SS:TdxSpreadSheet; Worksheet:integer; Row:integer);
 procedure SetMergeRowAutoBestHaight(SS:TdxSpreadSheet; Worksheet:integer; Row:integer; Column_From:integer; ColumnTo:integer; var set_row_Haight:Integer);

 procedure GroupRows(SS:TdxSpreadSheet; Worksheet:integer; AFirstRow, ALastRow: Integer);
 procedure GroupExpandAll(SS:TdxSpreadSheet; Worksheet:integer);

 procedure DrawCell(SS:TdxSpreadSheet;
                    CellText:String; //То, что хотим показать
                    Worksheet:integer; //Страница
                    n_row:integer;  //Координаты строки
                    n_col:integer; //Координаты столбца

                    FormulaChecking:boolean = True; //Это формула?
                    HAlign: TdxSpreadSheetDataAlignHorz = ssahGeneral; //Как выравнивать по горизонтали
                    //(ssahGeneral, ssahLeft, ssahCenter, ssahRight, ssahFill, ssahJustify, ssahDistributed)
                    VAlign: TdxSpreadSheetDataAlignVert = ssavBottom; //Как выравнивать по вертикали
                    //(ssavTop, ssavCenter, ssavBottom, ssavJustify, ssavDistributed)
                    sFormat:Integer=$00; //Форматирование ячейки
                                 // $00 GENERAL
                                 // $01 0
                                 // $02 0.00
                                 // $03 #,##0
                                 // $04 #,##0.00
                                 // $05 $#,##0_);($#,##0)
                                 // $06 $#,##0_);[Red]($#,##0)
                                 // $07 $#,##0.00_);($#,##0.00)
                                 // $08 $#,##0.00_);[Red]($#,##0.00)
                                 // $09 0%
                                 // $0a 0.00%
                                 // $0b 0.00E+00
                                 // $0c # ?/?
                                 // $0d # ??/??
                                 // $0e m/d/yy
                                 // $0f d-mmm-yy
                                 // $10 d-mmm
                                 // $11 mmm-yy
                                 // $12 h:mm AM/PM
                                 // $13 h:mm:ss AM/PM
                                 // $14 h:mm
                                 // $15 h:mm:ss
                                 // $16 m/d/yy h:mm
                                 // $25 #,##0_);(#,##0)
                                 // $26 #,##0_);[Red](#,##0)
                                 // $27 #,##0.00_);(#,##0.00)
                                 // $28 #,##0.00_);[Red](#,##0.00)
                                 // $29 _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
                                 // $2a _($* #,##0_);_($* (#,##0);_($* "-"_);_
                                 // $2b _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
                                 // $2c _($* #,##0.00);_($* (#,##0.00);_($* "-"??_);_(@_)
                                 // $2d mm:ss
                                 // $2e [h]:mm:ss
                                 // $2f mm:ss.0
                                 // $30 ##0.0E+0
                                 // $31 @

                    dWordWrap:boolean=False; //Если True, то содержание ячейки переносится, если оно не вписывается в ячейку по горизонтали.
                    //в противном случае, содержимое ячейки распространяется на пространство, занимаемое соседней ячейки
                    dShrinkToFit:boolean=False; //Если True, то осуществляется подгонка ячейки по ширине

                    //Шрифт ячейки
                    FontName:TFontName='';
                    FontSize:integer=0;
                    FontStyle:TFontStyles=[]; // (fsBold, fsItalic, fsUnderline, fsStrikeOut);
                                          // например  ([], [fsBold], [fsItalic], [fsBold, fsUnderline], [fsBold, fsItalic, fsStrikeOut])
                    FontColor: TColor=0;

                    //Бордюры ячейки
                    LeftColor: TColor=0;
                    LeftStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault; //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                  // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                  //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)

                    TopColor: TColor=0;
                    TopStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault;  //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                  // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                  //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)

                    RightColor: TColor=0;
                    RightStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault; //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                   // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                   //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)


                    BottomColor: TColor=0;
                    BottomStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault;//(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                   // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                   //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)


                    //Красим саму ячейку
                    wBackgroundColor: TColor=0;
                    wForegroundColor: TColor=0;
                    wStyle: TdxSpreadSheetCellFillStyle=sscfsSolid // (sscfsSolid, sscfsGray75, sscfsGray50, sscfsGray25, sscfsGray12, sscfsGray6,
                                                                 // sscfsHorzStrip, sscfsVertStrip, sscfsRevDiagonalStrip, sscfsDiagonalStrip,
                                                                 // sscfsDiagCrossHatch, sscfsThickCrossHatch, sscfsThinHorzStrip, sscfsThinVertStrip,
                                                                 // sscfsThinRevDiagonalStrip, sscfsThinDiagonalStrip, sscfsThinDiagCrossHatch,
                                                                 // sscfsThinThickCrossHatch)
                    );

       procedure DrawCellNotPaint(SS:TdxSpreadSheet;
                        CellText:String; //То, что хотим показать
                        Worksheet:integer; //Страница
                        n_row:integer;  //Координаты строки
                        n_col:integer; //Координаты столбца

                        FormulaChecking:boolean = True; //Это формула?
                        HAlign: TdxSpreadSheetDataAlignHorz = ssahGeneral; //Как выравнивать по горизонтали
                        //(ssahGeneral, ssahLeft, ssahCenter, ssahRight, ssahFill, ssahJustify, ssahDistributed)
                        VAlign: TdxSpreadSheetDataAlignVert = ssavBottom; //Как выравнивать по вертикали
                        //(ssavTop, ssavCenter, ssavBottom, ssavJustify, ssavDistributed)
                        sFormat:Integer=$00; //Форматирование ячейки
                                     // $00 GENERAL
                                     // $01 0
                                     // $02 0.00
                                     // $03 #,##0
                                     // $04 #,##0.00
                                     // $05 $#,##0_);($#,##0)
                                     // $06 $#,##0_);[Red]($#,##0)
                                     // $07 $#,##0.00_);($#,##0.00)
                                     // $08 $#,##0.00_);[Red]($#,##0.00)
                                     // $09 0%
                                     // $0a 0.00%
                                     // $0b 0.00E+00
                                     // $0c # ?/?
                                     // $0d # ??/??
                                     // $0e m/d/yy
                                     // $0f d-mmm-yy
                                     // $10 d-mmm
                                     // $11 mmm-yy
                                     // $12 h:mm AM/PM
                                     // $13 h:mm:ss AM/PM
                                     // $14 h:mm
                                     // $15 h:mm:ss
                                     // $16 m/d/yy h:mm
                                     // $25 #,##0_);(#,##0)
                                     // $26 #,##0_);[Red](#,##0)
                                     // $27 #,##0.00_);(#,##0.00)
                                     // $28 #,##0.00_);[Red](#,##0.00)
                                     // $29 _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
                                     // $2a _($* #,##0_);_($* (#,##0);_($* "-"_);_
                                     // $2b _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
                                     // $2c _($* #,##0.00);_($* (#,##0.00);_($* "-"??_);_(@_)
                                     // $2d mm:ss
                                     // $2e [h]:mm:ss
                                     // $2f mm:ss.0
                                     // $30 ##0.0E+0
                                     // $31 @

                        dWordWrap:boolean=False; //Если True, то содержание ячейки переносится, если оно не вписывается в ячейку по горизонтали.
                        //в противном случае, содержимое ячейки распространяется на пространство, занимаемое соседней ячейки
                        dShrinkToFit:boolean=False; //Если True, то осуществляется подгонка ячейки по ширине

                        //Шрифт ячейки
                        FontName:TFontName='';
                        FontSize:integer=0;
                        FontStyle:TFontStyles=[]; // (fsBold, fsItalic, fsUnderline, fsStrikeOut);
                                              // например  ([], [fsBold], [fsItalic], [fsBold, fsUnderline], [fsBold, fsItalic, fsStrikeOut])
                        FontColor: TColor=0
                    );

       procedure DrawCellOnlyFont(SS:TdxSpreadSheet;
                                  Worksheet:integer; //Страница
                                  n_row:integer;  //Координаты строки
                                  n_col:integer; //Координаты столбца

                                  //Шрифт ячейки
                                  FontName:TFontName='';
                                  FontSize:integer=0;
                                  FontStyle:TFontStyles=[]; // (fsBold, fsItalic, fsUnderline, fsStrikeOut);
                                              // например  ([], [fsBold], [fsItalic], [fsBold, fsUnderline], [fsBold, fsItalic, fsStrikeOut])
                                  FontColor: TColor=0
                                  );

       procedure DrawCellOnlyBorder(SS:TdxSpreadSheet;
                                    Worksheet:integer; //Страница
                                    n_row:integer;  //Координаты строки
                                    n_col:integer; //Координаты столбца

                                    //Бордюры ячейки
                                    LeftColor: TColor=0;
                                    LeftStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault; //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                                          // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                                         //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)

                                    TopColor: TColor=0;
                                    TopStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault;  //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                                          // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                                         //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)

                                    RightColor: TColor=0;
                                    RightStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault; //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                                           // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                                          //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)


                                    BottomColor: TColor=0;
                                    BottomStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault;//(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                                           // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                                          //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)


                                    //Красим саму ячейку
                                    wBackgroundColor: TColor=0;
                                    wForegroundColor: TColor=0;
                                    wStyle: TdxSpreadSheetCellFillStyle=sscfsSolid // (sscfsSolid, sscfsGray75, sscfsGray50, sscfsGray25, sscfsGray12, sscfsGray6,
                                                                                  // sscfsHorzStrip, sscfsVertStrip, sscfsRevDiagonalStrip, sscfsDiagonalStrip,
                                                                                 // sscfsDiagCrossHatch, sscfsThickCrossHatch, sscfsThinHorzStrip, sscfsThinVertStrip,
                                                                                // sscfsThinRevDiagonalStrip, sscfsThinDiagonalStrip, sscfsThinDiagCrossHatch,
                                                                               // sscfsThinThickCrossHatch)
                                   );

procedure DrawCellMergeCells(SS:TdxSpreadSheet;
                             Worksheet:integer; //Страница
                             nLeft:integer=0;
                             nTop:integer=0;
                             nRight:integer=0;
                             nBottom:integer=0;

                             //Бордюры ячейки
                             LeftColor: TColor=0;
                             LeftStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault; //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                  // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                  //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)

                             TopColor: TColor=0;
                             TopStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault;  //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                             // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                             //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)

                             RightColor: TColor=0;
                             RightStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault; //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                   // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                   //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)


                             BottomColor: TColor=0;
                             BottomStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault;//(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                   // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                   //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)


                              is_Cell_text:boolean = False; //Бум ли что-то писать или null
                              CellText:String=''; //То, что хотим показать
                              n_row:integer=0;  //Координаты строки
                              n_col:integer=0; //Координаты столбца

                              FormulaChecking:boolean = True; //Это формула?
                              HAlign: TdxSpreadSheetDataAlignHorz = ssahGeneral; //Как выравнивать по горизонтали
                                              //(ssahGeneral, ssahLeft, ssahCenter, ssahRight, ssahFill, ssahJustify, ssahDistributed)
                              VAlign: TdxSpreadSheetDataAlignVert = ssavBottom; //Как выравнивать по вертикали
                                              //(ssavTop, ssavCenter, ssavBottom, ssavJustify, ssavDistributed)
                              sFormat:Integer=$00; //Форматирование ячейки
                                                // $00 GENERAL
                                                // $01 0
                                                // $02 0.00
                                                // $03 #,##0
                                                // $04 #,##0.00
                                                // $05 $#,##0_);($#,##0)
                                                // $06 $#,##0_);[Red]($#,##0)
                                                // $07 $#,##0.00_);($#,##0.00)
                                                // $08 $#,##0.00_);[Red]($#,##0.00)
                                                // $09 0%
                                                // $0a 0.00%
                                                // $0b 0.00E+00
                                                // $0c # ?/?
                                                // $0d # ??/??
                                                // $0e m/d/yy
                                                // $0f d-mmm-yy
                                                // $10 d-mmm
                                                // $11 mmm-yy
                                                // $12 h:mm AM/PM
                                                // $13 h:mm:ss AM/PM
                                                // $14 h:mm
                                                // $15 h:mm:ss
                                                // $16 m/d/yy h:mm
                                                // $25 #,##0_);(#,##0)
                                                // $26 #,##0_);[Red](#,##0)
                                                // $27 #,##0.00_);(#,##0.00)
                                                // $28 #,##0.00_);[Red](#,##0.00)
                                                // $29 _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
                                                // $2a _($* #,##0_);_($* (#,##0);_($* "-"_);_
                                                // $2b _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
                                                // $2c _($* #,##0.00);_($* (#,##0.00);_($* "-"??_);_(@_)
                                                // $2d mm:ss
                                                // $2e [h]:mm:ss
                                                // $2f mm:ss.0
                                                // $30 ##0.0E+0
                                                // $31 @

                                                 dWordWrap:boolean=False; //Если True, то содержание ячейки переносится, если оно не вписывается в ячейку по горизонтали.
                                               //в противном случае, содержимое ячейки распространяется на пространство, занимаемое соседней ячейки
                                                 dShrinkToFit:boolean=False; //Если True, то осуществляется подгонка ячейки по ширине

                                               //Шрифт ячейки
                                                 FontName:TFontName='';
                                                 FontSize:integer=0;
                                                 FontStyle:TFontStyles=[]; // (fsBold, fsItalic, fsUnderline, fsStrikeOut);
                                               // например  ([], [fsBold], [fsItalic], [fsBold, fsUnderline], [fsBold, fsItalic, fsStrikeOut])
                                                 FontColor: TColor=0;

                                                //Красим саму ячейку
                                                 wBackgroundColor: TColor=0;
                                                 wForegroundColor: TColor=0;
                                                 wStyle: TdxSpreadSheetCellFillStyle=sscfsSolid; // (sscfsSolid, sscfsGray75, sscfsGray50, sscfsGray25, sscfsGray12, sscfsGray6,
                                                                 // sscfsHorzStrip, sscfsVertStrip, sscfsRevDiagonalStrip, sscfsDiagonalStrip,
                                                                 // sscfsDiagCrossHatch, sscfsThickCrossHatch, sscfsThinHorzStrip, sscfsThinVertStrip,
                                                                 // sscfsThinRevDiagonalStrip, sscfsThinDiagonalStrip, sscfsThinDiagCrossHatch,
                                                                 // sscfsThinThickCrossHatch)

                                                 is_auto_Haight_row:boolean = False
                                                 );


procedure DrawCellOnlyColor(SS:TdxSpreadSheet;
                            Worksheet:integer; //Страница
                            n_row:integer;  //Координаты строки
                            n_col:integer; //Координаты столбца
                                          //Красим саму ячейку
                            wBackgroundColor: TColor=0;
                            wForegroundColor: TColor=0;
                            wStyle: TdxSpreadSheetCellFillStyle=sscfsSolid // (sscfsSolid, sscfsGray75, sscfsGray50, sscfsGray25, sscfsGray12, sscfsGray6,
                                                                           // sscfsHorzStrip, sscfsVertStrip, sscfsRevDiagonalStrip, sscfsDiagonalStrip,
                                                                           // sscfsDiagCrossHatch, sscfsThickCrossHatch, sscfsThinHorzStrip, sscfsThinVertStrip,
                                                                           // sscfsThinRevDiagonalStrip, sscfsThinDiagonalStrip, sscfsThinDiagCrossHatch,
                                                                          // sscfsThinThickCrossHatch)
                                 );


procedure DrawCellOnlyValue(SS:TdxSpreadSheet;
                            Worksheet:integer; //Страница
                            n_row:integer;  //Координаты строки
                            n_col:integer; //Координаты столбца
                            CellText:String; //То, что хотим показать
                            FormulaChecking:boolean = True
                                 );




procedure CreateHyperlink(SS:TdxSpreadSheet; Worksheet:integer; Row:integer; Column:integer; ScreenTip: string; Value:variant);
function Get_Cell(SS:TdxSpreadSheet; Worksheet, Row:integer; Column:integer):TdxSpreadSheetCell;
procedure FreezeRows(SS:TdxSpreadSheet; Worksheet:integer; Row:integer);
procedure FreezeColumns(SS:TdxSpreadSheet; Worksheet:integer; Column:integer);
procedure SetHorAllignCenter(SS:TdxSpreadSheet; Worksheet:integer; Row:integer; Column:integer);

procedure SetDefaultRowHeight(SS:TdxSpreadSheet; Worksheet:integer; RowHeight:integer);
procedure SetDefaultColumnWidth(SS:TdxSpreadSheet; Worksheet:integer; ColumnWidth:integer);
procedure DrawRect(SS:TdxSpreadSheet; Worksheet, nLeft,nTop,nRight,nBottom:integer; RStyle:TdxSpreadSheetCellBorderStyle; is_merge:boolean=False);

function  GetRowHaight(SS:TdxSpreadSheet; Worksheet:integer; Row:integer):integer;
procedure CopyHyperlinks_to_Temp(SS:TdxSpreadSheet; Worksheet:integer);
procedure CopyHyperlinks_from_Temp(SS:TdxSpreadSheet; Worksheet:integer);
procedure RecalcFormulsWorksheet(SS:TdxSpreadSheet);
procedure ExelSaveToFile(SS:TdxSpreadSheet; FileName:string);
procedure ClearCells_(SS:TdxSpreadSheet; Worksheet:integer; col_left:integer; col_right:integer; row_up:integer; row_down:integer);
procedure SetFocusedCell_(SS:TdxSpreadSheet; Worksheet:integer; col:integer; row:integer);


implementation


function AddWorksheet(SS:TdxSpreadSheet; new_caption:string):integer;
var
 t_TdxSpreadSheetCustomView:TdxSpreadSheetCustomView;
begin
  t_TdxSpreadSheetCustomView:=SS.AddSheet(new_caption);
  Result:=t_TdxSpreadSheetCustomView.Index;
end;

procedure RenameWorksheet(SS:TdxSpreadSheet; Worksheet:integer; new_caption:string);
begin
  SS.Sheets[Worksheet].Caption := new_caption;
end;

procedure MergeCells(SS:TdxSpreadSheet; Worksheet,nLeft,nTop,nRight,nBottom:integer);
var
ARect: TRect;
begin
  with ARect do
  begin
    Left := nLeft-1;
    Top := nTop-1;
    Right := nRight-1;
    Bottom := nBottom-1;
  end;
  TdxSpreadSheetTableView(SS.Sheets[Worksheet]).MergedCells.Add(ARect);
end;

procedure SetColumnWidth(SS:TdxSpreadSheet; Worksheet:integer; Column:integer; Width:Integer);
var
ATableView: TdxSpreadSheetTableView;
AColumn: TdxSpreadSheetTableColumn;
begin
 ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
 if(ATableView.Columns[Column-1] = nil) then
 begin
    ATableView.Columns.CreateItem(Column-1);
 end;
 AColumn := ATableView.Columns[Column-1];
 AColumn.Size:=Width;
 AColumn:=nil;
end;

procedure SetRowHaight(SS:TdxSpreadSheet; Worksheet:integer; Row:integer; Haight:Integer);
var
ATableView: TdxSpreadSheetTableView;
ARow: TdxSpreadSheetTableRow;
begin
   ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
   if(ATableView.Rows[Row-1] = nil) then
   begin
    ATableView.Rows.CreateItem(Row-1);
   end;
   ARow := ATableView.Rows[Row-1];
   ARow.Size:=Haight;
   ARow:=nil;
end;

procedure SetColumnAutoBestWidth(SS:TdxSpreadSheet; Worksheet:integer; Column:integer);
var
ATableView: TdxSpreadSheetTableView;
AColumn: TdxSpreadSheetTableColumn;
begin
 ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
 if(ATableView.Columns[Column-1] = nil) then
 begin
    ATableView.Columns.CreateItem(Column-1);
 end;
 AColumn := ATableView.Columns[Column-1];
 AColumn.ApplyBestFit;
 AColumn:=nil;
end;

procedure SetRowAutoBestHaight(SS:TdxSpreadSheet; Worksheet:integer; Row:integer);
var
ATableView: TdxSpreadSheetTableView;
ARow: TdxSpreadSheetTableRow;
begin
   ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
   if(ATableView.Rows[Row-1] = nil) then
   begin
    ATableView.Rows.CreateItem(Row-1);
   end;
   ARow := ATableView.Rows[Row-1];
   ARow.ApplyBestFit;
   ARow:=nil;
end;

procedure SetMergeRowAutoBestHaight(SS:TdxSpreadSheet; Worksheet:integer; Row:integer; Column_From:integer; ColumnTo:integer; var set_row_Haight:Integer);
var
ATableView: TdxSpreadSheetTableView;
ARow: TdxSpreadSheetTableRow;
AColumn: TdxSpreadSheetTableColumn;
n:integer;
size_all_original:integer;
size_all_colmns:Integer;
size_row_original:integer;
_DefaultRowHeight:integer;
Count_row:Extended;
begin
  ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
  _DefaultRowHeight:=ATableView.Options.DefaultRowHeight;

  if(ATableView.Rows[Row-1] = nil) then
  begin
   ATableView.Rows.CreateItem(Row-1);
  end;
  ARow := ATableView.Rows[Row-1];
  size_row_original:=ARow.Size;
  ARow:=nil;


  size_all_colmns:=0;
  for n:=Column_From-1 to ColumnTo-1 do
  begin
     if(ATableView.Columns[n] = nil) then
     begin
       ATableView.Columns.CreateItem(n);
     end;
     AColumn := ATableView.Columns[n];
     size_all_colmns:=size_all_colmns+AColumn.Size;
     AColumn:=nil;
  end;

  if(ATableView.Columns[Column_From-1] = nil) then
  begin
    ATableView.Columns.CreateItem(Column_From-1);
  end;
  AColumn := ATableView.Columns[Column_From-1];
  size_all_original:=AColumn.Size;
  AColumn.Size:=size_all_colmns;
  AColumn:=nil;

  SetRowAutoBestHaight(SS, Worksheet, Row);

  if(ATableView.Rows[Row-1] = nil) then
  begin
   ATableView.Rows.CreateItem(Row-1);
  end;
  ARow := ATableView.Rows[Row-1];

  Count_row:=ARow.Size/_DefaultRowHeight;

  if ARow.Size>=size_row_original then
  begin
//    set_row_Haight:=ARow.Size
    set_row_Haight:=Ceil(Count_row*size_row_original);
  end
  else
  begin
    set_row_Haight:=size_row_original;
  end;
  ARow:=nil;


  if(ATableView.Columns[Column_From-1] = nil) then
  begin
    ATableView.Columns.CreateItem(Column_From-1);
  end;
  AColumn := ATableView.Columns[Column_From-1];
  AColumn.Size:=size_all_original;
  AColumn:=nil;
end;

procedure GroupRows(SS:TdxSpreadSheet; Worksheet:integer; AFirstRow, ALastRow: Integer);
var
  ATableView: TdxSpreadSheetTableView;
begin
  ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
  ATableView.Rows.Groups.Add(AFirstRow-1, ALastRow-1);
end;

procedure GroupExpandAll(SS:TdxSpreadSheet; Worksheet:integer);
var
  ATableView: TdxSpreadSheetTableView;
  i:integer;
begin
  ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
  for i := 0 to ATableView.Rows.Groups.Count-1 do
  begin
   ATableView.Rows.Groups.Items[i].Expanded:=False;
  end;
end;

procedure DrawCell(SS:TdxSpreadSheet;
                    CellText:String; //То, что хотим показать
                    Worksheet:integer; //Страница
                    n_row:integer;  //Координаты строки
                    n_col:integer; //Координаты столбца

                    FormulaChecking:boolean = True; //Это формула?
                    HAlign: TdxSpreadSheetDataAlignHorz = ssahGeneral; //Как выравнивать по горизонтали
                    //(ssahGeneral, ssahLeft, ssahCenter, ssahRight, ssahFill, ssahJustify, ssahDistributed)
                    VAlign: TdxSpreadSheetDataAlignVert = ssavBottom; //Как выравнивать по вертикали
                    //(ssavTop, ssavCenter, ssavBottom, ssavJustify, ssavDistributed)
                    sFormat:Integer=$00; //Форматирование ячейки
                                 // $00 GENERAL
                                 // $01 0
                                 // $02 0.00
                                 // $03 #,##0
                                 // $04 #,##0.00
                                 // $05 $#,##0_);($#,##0)
                                 // $06 $#,##0_);[Red]($#,##0)
                                 // $07 $#,##0.00_);($#,##0.00)
                                 // $08 $#,##0.00_);[Red]($#,##0.00)
                                 // $09 0%
                                 // $0a 0.00%
                                 // $0b 0.00E+00
                                 // $0c # ?/?
                                 // $0d # ??/??
                                 // $0e m/d/yy
                                 // $0f d-mmm-yy
                                 // $10 d-mmm
                                 // $11 mmm-yy
                                 // $12 h:mm AM/PM
                                 // $13 h:mm:ss AM/PM
                                 // $14 h:mm
                                 // $15 h:mm:ss
                                 // $16 m/d/yy h:mm
                                 // $25 #,##0_);(#,##0)
                                 // $26 #,##0_);[Red](#,##0)
                                 // $27 #,##0.00_);(#,##0.00)
                                 // $28 #,##0.00_);[Red](#,##0.00)
                                 // $29 _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
                                 // $2a _($* #,##0_);_($* (#,##0);_($* "-"_);_
                                 // $2b _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
                                 // $2c _($* #,##0.00);_($* (#,##0.00);_($* "-"??_);_(@_)
                                 // $2d mm:ss
                                 // $2e [h]:mm:ss
                                 // $2f mm:ss.0
                                 // $30 ##0.0E+0
                                 // $31 @

                    dWordWrap:boolean=False; //Если True, то содержание ячейки переносится, если оно не вписывается в ячейку по горизонтали.
                    //в противном случае, содержимое ячейки распространяется на пространство, занимаемое соседней ячейки
                    dShrinkToFit:boolean=False; //Если True, то осуществляется подгонка ячейки по ширине

                    //Шрифт ячейки
                    FontName:TFontName='';
                    FontSize:integer=0;
                    FontStyle:TFontStyles=[]; // (fsBold, fsItalic, fsUnderline, fsStrikeOut);
                                          // например  ([], [fsBold], [fsItalic], [fsBold, fsUnderline], [fsBold, fsItalic, fsStrikeOut])
                    FontColor: TColor=0;

                    //Бордюры ячейки
                    LeftColor: TColor=0;
                    LeftStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault; //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                  // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                  //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)

                    TopColor: TColor=0;
                    TopStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault;  //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                  // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                  //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)

                    RightColor: TColor=0;
                    RightStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault; //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                   // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                   //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)


                    BottomColor: TColor=0;
                    BottomStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault;//(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                   // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                   //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)


                    //Красим саму ячейку
                    wBackgroundColor: TColor=0;
                    wForegroundColor: TColor=0;
                    wStyle: TdxSpreadSheetCellFillStyle=sscfsSolid // (sscfsSolid, sscfsGray75, sscfsGray50, sscfsGray25, sscfsGray12, sscfsGray6,
                                                                 // sscfsHorzStrip, sscfsVertStrip, sscfsRevDiagonalStrip, sscfsDiagonalStrip,
                                                                 // sscfsDiagCrossHatch, sscfsThickCrossHatch, sscfsThinHorzStrip, sscfsThinVertStrip,
                                                                 // sscfsThinRevDiagonalStrip, sscfsThinDiagonalStrip, sscfsThinDiagCrossHatch,
                                                                 // sscfsThinThickCrossHatch)
                    );
 var
   Cell: TdxSpreadSheetCell;
   ATableView: TdxSpreadSheetTableView;
 begin
   ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
   Cell:=ATableView.CreateCell(n_row-1, n_col-1);
   Cell.Style.AlignHorz:=HAlign;
   Cell.Style.AlignVert:=VAlign;
   Cell.Style.DataFormat.FormatCodeID:=sFormat;
   Cell.Style.WordWrap:=dWordWrap;
   Cell.Style.ShrinkToFit:=dShrinkToFit;

   with Cell.Style.Font do
   begin
     if FontName<>'' then
     Name := FontName;

     if FontSize<>0 then
     Size := FontSize;

     if FontStyle<>[] then
     Style := FontStyle;

     if FontColor<>0 then
     Color:=FontColor;
   end;

   if LeftColor<>0 then
   Cell.Style.Borders[bLeft].Color := LeftColor;
   Cell.Style.Borders[bLeft].Style := LeftStyle;
   if TopColor<>0 then
   Cell.Style.Borders[bTop].Color := TopColor;
   Cell.Style.Borders[bTop].Style := TopStyle;
   if RightColor<>0 then
   Cell.Style.Borders[bRight].Color := RightColor;
   Cell.Style.Borders[bRight].Style := RightStyle;
   if BottomColor<>0 then
   Cell.Style.Borders[bBottom].Color := BottomColor;
   Cell.Style.Borders[bBottom].Style := BottomStyle;

   with Cell.Style.Brush do
   begin
    if wBackgroundColor<>0 then
    BackgroundColor := wBackgroundColor;
    if wForegroundColor<>0 then
    ForegroundColor := wForegroundColor;
    Style := wStyle;
   end;

   Cell.SetText(CellText, FormulaChecking);
 end;

procedure DrawCellNotPaint(SS:TdxSpreadSheet;
                        CellText:String; //То, что хотим показать
                        Worksheet:integer; //Страница
                        n_row:integer;  //Координаты строки
                        n_col:integer; //Координаты столбца

                        FormulaChecking:boolean = True; //Это формула?
                        HAlign: TdxSpreadSheetDataAlignHorz = ssahGeneral; //Как выравнивать по горизонтали
                        //(ssahGeneral, ssahLeft, ssahCenter, ssahRight, ssahFill, ssahJustify, ssahDistributed)
                        VAlign: TdxSpreadSheetDataAlignVert = ssavBottom; //Как выравнивать по вертикали
                        //(ssavTop, ssavCenter, ssavBottom, ssavJustify, ssavDistributed)
                        sFormat:Integer=$00; //Форматирование ячейки
                                     // $00 GENERAL
                                     // $01 0
                                     // $02 0.00
                                     // $03 #,##0
                                     // $04 #,##0.00
                                     // $05 $#,##0_);($#,##0)
                                     // $06 $#,##0_);[Red]($#,##0)
                                     // $07 $#,##0.00_);($#,##0.00)
                                     // $08 $#,##0.00_);[Red]($#,##0.00)
                                     // $09 0%
                                     // $0a 0.00%
                                     // $0b 0.00E+00
                                     // $0c # ?/?
                                     // $0d # ??/??
                                     // $0e m/d/yy
                                     // $0f d-mmm-yy
                                     // $10 d-mmm
                                     // $11 mmm-yy
                                     // $12 h:mm AM/PM
                                     // $13 h:mm:ss AM/PM
                                     // $14 h:mm
                                     // $15 h:mm:ss
                                     // $16 m/d/yy h:mm
                                     // $25 #,##0_);(#,##0)
                                     // $26 #,##0_);[Red](#,##0)
                                     // $27 #,##0.00_);(#,##0.00)
                                     // $28 #,##0.00_);[Red](#,##0.00)
                                     // $29 _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
                                     // $2a _($* #,##0_);_($* (#,##0);_($* "-"_);_
                                     // $2b _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
                                     // $2c _($* #,##0.00);_($* (#,##0.00);_($* "-"??_);_(@_)
                                     // $2d mm:ss
                                     // $2e [h]:mm:ss
                                     // $2f mm:ss.0
                                     // $30 ##0.0E+0
                                     // $31 @

                        dWordWrap:boolean=False; //Если True, то содержание ячейки переносится, если оно не вписывается в ячейку по горизонтали.
                        //в противном случае, содержимое ячейки распространяется на пространство, занимаемое соседней ячейки
                        dShrinkToFit:boolean=False; //Если True, то осуществляется подгонка ячейки по ширине

                        //Шрифт ячейки
                        FontName:TFontName='';
                        FontSize:integer=0;
                        FontStyle:TFontStyles=[]; // (fsBold, fsItalic, fsUnderline, fsStrikeOut);
                                              // например  ([], [fsBold], [fsItalic], [fsBold, fsUnderline], [fsBold, fsItalic, fsStrikeOut])
                        FontColor: TColor=0
                    );
 var
   Cell: TdxSpreadSheetCell;
   ATableView: TdxSpreadSheetTableView;
 begin
   ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
   Cell:=ATableView.CreateCell(n_row-1, n_col-1);

   Cell.Style.AlignHorz:=HAlign;
   Cell.Style.AlignVert:=VAlign;

   Cell.Style.DataFormat.FormatCodeID:=sFormat;
   Cell.Style.WordWrap:=dWordWrap;
   Cell.Style.ShrinkToFit:=dShrinkToFit;

   with Cell.Style.Font do
   begin
     if FontName<>'' then
     Name := FontName;

     if FontSize<>0 then
     Size := FontSize;

     if FontStyle<>[] then
     Style := FontStyle;

     if FontColor<>0 then
     Color:=FontColor;
   end;

   Cell.SetText(CellText, FormulaChecking);
 end;


procedure DrawCellOnlyFont(SS:TdxSpreadSheet;
                                  Worksheet:integer; //Страница
                                  n_row:integer;  //Координаты строки
                                  n_col:integer; //Координаты столбца

                                  //Шрифт ячейки
                                  FontName:TFontName='';
                                  FontSize:integer=0;
                                  FontStyle:TFontStyles=[]; // (fsBold, fsItalic, fsUnderline, fsStrikeOut);
                                              // например  ([], [fsBold], [fsItalic], [fsBold, fsUnderline], [fsBold, fsItalic, fsStrikeOut])
                                  FontColor: TColor=0
                                  );
 var
   Cell: TdxSpreadSheetCell;
   ATableView: TdxSpreadSheetTableView;
 begin
   ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
   Cell:=ATableView.CreateCell(n_row-1, n_col-1);

   with Cell.Style.Font do
   begin
     if FontName<>'' then
     Name := FontName;

     if FontSize<>0 then
     Size := FontSize;

     if FontStyle<>[] then
     Style := FontStyle;

     if FontColor<>0 then
     Color:=FontColor;
   end;
 end;


 procedure DrawCellOnlyBorder(SS:TdxSpreadSheet;
                                    Worksheet:integer; //Страница
                                    n_row:integer;  //Координаты строки
                                    n_col:integer; //Координаты столбца

                                    //Бордюры ячейки
                                    LeftColor: TColor=0;
                                    LeftStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault; //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                                          // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                                         //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)

                                    TopColor: TColor=0;
                                    TopStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault;  //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                                          // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                                         //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)

                                    RightColor: TColor=0;
                                    RightStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault; //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                                           // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                                          //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)


                                    BottomColor: TColor=0;
                                    BottomStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault;//(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                                           // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                                          //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)


                                    //Красим саму ячейку
                                    wBackgroundColor: TColor=0;
                                    wForegroundColor: TColor=0;
                                    wStyle: TdxSpreadSheetCellFillStyle=sscfsSolid // (sscfsSolid, sscfsGray75, sscfsGray50, sscfsGray25, sscfsGray12, sscfsGray6,
                                                                                  // sscfsHorzStrip, sscfsVertStrip, sscfsRevDiagonalStrip, sscfsDiagonalStrip,
                                                                                 // sscfsDiagCrossHatch, sscfsThickCrossHatch, sscfsThinHorzStrip, sscfsThinVertStrip,
                                                                                // sscfsThinRevDiagonalStrip, sscfsThinDiagonalStrip, sscfsThinDiagCrossHatch,
                                                                               // sscfsThinThickCrossHatch)
                                   );
 var
   Cell: TdxSpreadSheetCell;
   ATableView: TdxSpreadSheetTableView;
 begin
   ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
   Cell:=ATableView.CreateCell(n_row-1, n_col-1);

   if LeftColor<>0 then
   Cell.Style.Borders[bLeft].Color := LeftColor;
   Cell.Style.Borders[bLeft].Style := LeftStyle;
   if TopColor<>0 then
   Cell.Style.Borders[bTop].Color := TopColor;
   Cell.Style.Borders[bTop].Style := TopStyle;
   if RightColor<>0 then
   Cell.Style.Borders[bRight].Color := RightColor;
   Cell.Style.Borders[bRight].Style := RightStyle;
   if BottomColor<>0 then
   Cell.Style.Borders[bBottom].Color := BottomColor;
   Cell.Style.Borders[bBottom].Style := BottomStyle;

   with Cell.Style.Brush do
   begin
    if wBackgroundColor<>0 then
    BackgroundColor := wBackgroundColor;
    if wForegroundColor<>0 then
    ForegroundColor := wForegroundColor;
    Style := wStyle;
   end;
 end;


procedure DrawCellMergeCells(SS:TdxSpreadSheet;
                             Worksheet:integer; //Страница
                             nLeft:integer=0;
                             nTop:integer=0;
                             nRight:integer=0;
                             nBottom:integer=0;

                             //Бордюры ячейки
                             LeftColor: TColor=0;
                             LeftStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault; //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                  // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                  //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)

                             TopColor: TColor=0;
                             TopStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault;  //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                             // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                             //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)

                             RightColor: TColor=0;
                             RightStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault; //(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                   // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                   //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)


                             BottomColor: TColor=0;
                             BottomStyle: TdxSpreadSheetCellBorderStyle=sscbsDefault;//(sscbsDefault, sscbsHair, sscbsDotted, sscbsDashDotDot, sscbsDashDot, sscbsDashed,
                                                                   // sscbsThin, sscbsMediumDashDotDot, sscbsSlantedDashDot, sscbsMediumDashDot,
                                                                   //sscbsMediumDashed, sscbsMedium, sscbsThick, sscbsDouble, sscbsNone)


                              is_Cell_text:boolean = False; //Бум ли что-то писать или null
                              CellText:String=''; //То, что хотим показать
                              n_row:integer=0;  //Координаты строки
                              n_col:integer=0; //Координаты столбца

                              FormulaChecking:boolean = True; //Это формула?
                              HAlign: TdxSpreadSheetDataAlignHorz = ssahGeneral; //Как выравнивать по горизонтали
                                              //(ssahGeneral, ssahLeft, ssahCenter, ssahRight, ssahFill, ssahJustify, ssahDistributed)
                              VAlign: TdxSpreadSheetDataAlignVert = ssavBottom; //Как выравнивать по вертикали
                                              //(ssavTop, ssavCenter, ssavBottom, ssavJustify, ssavDistributed)
                              sFormat:Integer=$00; //Форматирование ячейки
                                                // $00 GENERAL
                                                // $01 0
                                                // $02 0.00
                                                // $03 #,##0
                                                // $04 #,##0.00
                                                // $05 $#,##0_);($#,##0)
                                                // $06 $#,##0_);[Red]($#,##0)
                                                // $07 $#,##0.00_);($#,##0.00)
                                                // $08 $#,##0.00_);[Red]($#,##0.00)
                                                // $09 0%
                                                // $0a 0.00%
                                                // $0b 0.00E+00
                                                // $0c # ?/?
                                                // $0d # ??/??
                                                // $0e m/d/yy
                                                // $0f d-mmm-yy
                                                // $10 d-mmm
                                                // $11 mmm-yy
                                                // $12 h:mm AM/PM
                                                // $13 h:mm:ss AM/PM
                                                // $14 h:mm
                                                // $15 h:mm:ss
                                                // $16 m/d/yy h:mm
                                                // $25 #,##0_);(#,##0)
                                                // $26 #,##0_);[Red](#,##0)
                                                // $27 #,##0.00_);(#,##0.00)
                                                // $28 #,##0.00_);[Red](#,##0.00)
                                                // $29 _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
                                                // $2a _($* #,##0_);_($* (#,##0);_($* "-"_);_
                                                // $2b _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
                                                // $2c _($* #,##0.00);_($* (#,##0.00);_($* "-"??_);_(@_)
                                                // $2d mm:ss
                                                // $2e [h]:mm:ss
                                                // $2f mm:ss.0
                                                // $30 ##0.0E+0
                                                // $31 @

                                                 dWordWrap:boolean=False; //Если True, то содержание ячейки переносится, если оно не вписывается в ячейку по горизонтали.
                                               //в противном случае, содержимое ячейки распространяется на пространство, занимаемое соседней ячейки
                                                 dShrinkToFit:boolean=False; //Если True, то осуществляется подгонка ячейки по ширине

                                               //Шрифт ячейки
                                                 FontName:TFontName='';
                                                 FontSize:integer=0;
                                                 FontStyle:TFontStyles=[]; // (fsBold, fsItalic, fsUnderline, fsStrikeOut);
                                               // например  ([], [fsBold], [fsItalic], [fsBold, fsUnderline], [fsBold, fsItalic, fsStrikeOut])
                                                 FontColor: TColor=0;

                                                //Красим саму ячейку
                                                 wBackgroundColor: TColor=0;
                                                 wForegroundColor: TColor=0;
                                                 wStyle: TdxSpreadSheetCellFillStyle=sscfsSolid; // (sscfsSolid, sscfsGray75, sscfsGray50, sscfsGray25, sscfsGray12, sscfsGray6,
                                                                 // sscfsHorzStrip, sscfsVertStrip, sscfsRevDiagonalStrip, sscfsDiagonalStrip,
                                                                 // sscfsDiagCrossHatch, sscfsThickCrossHatch, sscfsThinHorzStrip, sscfsThinVertStrip,
                                                                 // sscfsThinRevDiagonalStrip, sscfsThinDiagonalStrip, sscfsThinDiagCrossHatch,
                                                                 // sscfsThinThickCrossHatch)

                                                 is_auto_Haight_row:boolean = False
                                                 );
var
 nrow, ncol, temp_integer:integer;
begin
   for nrow:= nTop to nBottom do
   for ncol:= nLeft to nRight do
   DrawCellOnlyBorder(SS, Worksheet, nrow, ncol, LeftColor, LeftStyle, TopColor, TopStyle, RightColor, RightStyle, BottomColor, BottomStyle);

   if is_auto_Haight_row = False then
   MergeCells(SS, Worksheet, nLeft,nTop,nRight,nBottom);

   if is_Cell_text = True then
   DrawCell(SS, CellText, Worksheet, n_row, n_col, FormulaChecking, HAlign, VAlign, sFormat, dWordWrap, dShrinkToFit, FontName, FontSize, FontStyle, FontColor,
   LeftColor, LeftStyle, TopColor, TopStyle, RightColor, RightStyle, BottomColor, BottomStyle, wBackgroundColor, wForegroundColor, wStyle);

   if is_auto_Haight_row = True then
   begin
     SetMergeRowAutoBestHaight(SS, Worksheet, n_row, nLeft,nRight, temp_integer);
     MergeCells(SS, Worksheet, nLeft,nTop,nRight,nBottom);
     SetRowHaight(SS, Worksheet, n_row, temp_integer);
   end;
end;


procedure DrawCellOnlyColor(SS:TdxSpreadSheet;
                            Worksheet:integer; //Страница
                            n_row:integer;  //Координаты строки
                            n_col:integer; //Координаты столбца
                                          //Красим саму ячейку
                            wBackgroundColor: TColor=0;
                            wForegroundColor: TColor=0;
                            wStyle: TdxSpreadSheetCellFillStyle=sscfsSolid // (sscfsSolid, sscfsGray75, sscfsGray50, sscfsGray25, sscfsGray12, sscfsGray6,
                                                                           // sscfsHorzStrip, sscfsVertStrip, sscfsRevDiagonalStrip, sscfsDiagonalStrip,
                                                                           // sscfsDiagCrossHatch, sscfsThickCrossHatch, sscfsThinHorzStrip, sscfsThinVertStrip,
                                                                           // sscfsThinRevDiagonalStrip, sscfsThinDiagonalStrip, sscfsThinDiagCrossHatch,
                                                                          // sscfsThinThickCrossHatch)
                                 );
 var
   Cell: TdxSpreadSheetCell;
   ATableView: TdxSpreadSheetTableView;
 begin
   ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
   Cell:=ATableView.CreateCell(n_row-1, n_col-1);

   with Cell.Style.Brush do
   begin
    if wBackgroundColor<>0 then
    BackgroundColor := wBackgroundColor;
    if wForegroundColor<>0 then
    ForegroundColor := wForegroundColor;
    Style := wStyle;
   end;
 end;


procedure DrawCellOnlyValue(SS:TdxSpreadSheet;
                            Worksheet:integer; //Страница
                            n_row:integer;  //Координаты строки
                            n_col:integer; //Координаты столбца
                            CellText:String; //То, что хотим показать
                            FormulaChecking:boolean = True //Это формула?
                            );
 var
   Cell: TdxSpreadSheetCell;
   ATableView: TdxSpreadSheetTableView;
begin
   ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
   Cell:=ATableView.CreateCell(n_row-1, n_col-1);
   Cell.SetText(CellText, FormulaChecking);
end;






procedure CreateHyperlink(SS:TdxSpreadSheet; Worksheet:integer; Row:integer; Column:integer; ScreenTip: string; Value:variant);
var
  ATableView: TdxSpreadSheetTableView;
//  AShapeContainer: TdxSpreadSheetShapeContainer;
  APictureContainer: TdxSpreadSheetPictureContainer;
  row_l:integer;
  column_l:integer;
begin
  ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);

  //  AShapeContainer := ATableView.Containers.Add(TdxSpreadSheetShapeContainer) as TdxSpreadSheetShapeContainer;
  //  AShapeContainer.Hyperlink := ATableView.Hyperlinks.Add(Row-1, Column-1);
  //  AShapeContainer.Hyperlink.ScreenTip:=ScreenTip;
  //  AShapeContainer.Hyperlink.Value:='';

  APictureContainer := ATableView.Containers.Add(TdxSpreadSheetPictureContainer) as TdxSpreadSheetPictureContainer;
  APictureContainer.Hyperlink := ATableView.Hyperlinks.Add(Rect(Column-1, Row-1, Column-1, Row-1));
  APictureContainer.Hyperlink.Value:= '';
  APictureContainer.Hyperlink.ScreenTip:=ScreenTip;


  row_l:=Length(arr_hiperlinks);
  if row_l<>0 then
  column_l:=Length(arr_hiperlinks[0]) else
  column_l:=0;

  if ((row_l<Row) or (column_l<Column)) then
  setlength(arr_hiperlinks, Row, Column);

  arr_hiperlinks[Row-1, Column-1]:=Value;
end;

function Get_Cell(SS:TdxSpreadSheet; Worksheet, Row:integer; Column:integer):TdxSpreadSheetCell;
var
 ATableView: TdxSpreadSheetTableView;
begin
 ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
 Result:=ATableView.CreateCell(Row-1, Column-1);
end;

procedure FreezeRows(SS:TdxSpreadSheet; Worksheet:integer; Row:integer);
var
 ATableView: TdxSpreadSheetTableView;
begin
 ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
 ATableView.FrozenRow := Row-1;
end;

procedure FreezeColumns(SS:TdxSpreadSheet; Worksheet:integer; Column:integer);
var
 ATableView: TdxSpreadSheetTableView;
begin
 ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
 ATableView.FrozenColumn:=Column-1;
end;

procedure SetHorAllignCenter(SS:TdxSpreadSheet; Worksheet:integer; Row:integer; Column:integer);
var
 ATableView: TdxSpreadSheetTableView;
 Cell: TdxSpreadSheetCell;
begin
 ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
 Cell:=ATableView.CreateCell(Row-1, Column-1);
 Cell.Style.AlignHorz:=ssahCenter;
end;

procedure SetDefaultRowHeight(SS:TdxSpreadSheet; Worksheet:integer; RowHeight:integer);
var
 ATableView: TdxSpreadSheetTableView;
begin
 ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
 ATableView.Options.DefaultRowHeight:=RowHeight;
end;

procedure SetDefaultColumnWidth(SS:TdxSpreadSheet; Worksheet:integer; ColumnWidth:integer);
var
 ATableView: TdxSpreadSheetTableView;
begin
 ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
 ATableView.Options.DefaultColumnWidth:=ColumnWidth;
end;

procedure DrawRect(SS:TdxSpreadSheet; Worksheet, nLeft,nTop,nRight,nBottom:integer; RStyle:TdxSpreadSheetCellBorderStyle; is_merge:boolean=False);
var
   Cell: TdxSpreadSheetCell;
   ATableView: TdxSpreadSheetTableView;
   ncol, nrow:integer;
   ARect: TRect;
begin
   ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);

   for ncol:= nLeft to nRight do
     for nrow:= nTop to nBottom do
   begin
     Cell:=ATableView.CreateCell(nrow-1, ncol-1);

     if ncol = nLeft then Cell.Style.Borders[bLeft].Style := RStyle;
     if ncol = nRight then Cell.Style.Borders[bRight].Style := RStyle;

     if nrow = nTop then Cell.Style.Borders[bTop].Style := RStyle;
     if nrow = nBottom then Cell.Style.Borders[bBottom].Style := RStyle;
   end;

   if is_merge = True then
   begin
      with ARect do
      begin
        Left := nLeft-1;
        Top := nTop-1;
        Right := nRight-1;
        Bottom := nBottom-1;
      end;
      TdxSpreadSheetTableView(SS.Sheets[Worksheet]).MergedCells.Add(ARect);
   end;
end;

function  GetRowHaight(SS:TdxSpreadSheet; Worksheet:integer; Row:integer):integer;
var
ATableView: TdxSpreadSheetTableView;
ARow: TdxSpreadSheetTableRow;
begin
  ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);

  if(ATableView.Rows[Row-1] = nil) then
  begin
   ATableView.Rows.CreateItem(Row-1);
  end;
  ARow := ATableView.Rows[Row-1];
  Result:=ARow.Size;
  ARow:=nil;
end;

procedure CopyHyperlinks_to_Temp(SS:TdxSpreadSheet; Worksheet:integer);
var
  ATableView: TdxSpreadSheetTableView;
  temp_ATableView: TdxSpreadSheetTableView;
  temp_APictureContainer: TdxSpreadSheetPictureContainer;
  i, Current_row, Current_col:integer;
  FontName:TFontName;
  FontSize:integer;
  FontStyle:TFontStyles;
  FontColor: TColor;
begin
  ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
  if (ATableView.Hyperlinks.Count>0) then
  begin
    temp_SS:=TdxSpreadSheet.Create(nil);
    temp_SS.AddSheet(SS.Sheets[Worksheet].Caption);
    temp_ATableView := TdxSpreadSheetTableView(temp_SS.Sheets[Worksheet]);

    temp_SS.Sheets[Worksheet].BeginUpdate;
    for i:=0 to ATableView.Hyperlinks.Count-1 do
    begin
      temp_APictureContainer := temp_ATableView.Containers.Add(TdxSpreadSheetPictureContainer) as TdxSpreadSheetPictureContainer;
      temp_APictureContainer.Hyperlink := temp_ATableView.Hyperlinks.Add(ATableView.Hyperlinks[i].Area);
      temp_APictureContainer.Hyperlink.Value := ATableView.Hyperlinks[i].Value;
      temp_APictureContainer.Hyperlink.ScreenTip:=ATableView.Hyperlinks[i].ScreenTip;
    end;
    temp_SS.Sheets[Worksheet].EndUpdate;
    FreeAndNil(temp_APictureContainer);

    SS.Sheets[Worksheet].BeginUpdate;
    for i:=0 to ATableView.Hyperlinks.Count-1 do
    begin
     Current_row:=ATableView.Hyperlinks[0].Area.Top;
     Current_col:=ATableView.Hyperlinks[0].Area.Left;


     FontName:=ATableView.Cells[Current_row, Current_col].Style.Font.Name;
     FontSize:=ATableView.Cells[Current_row, Current_col].Style.Font.Size;
     FontStyle:=ATableView.Cells[Current_row, Current_col].Style.Font.Style;
     FontColor:=ATableView.Cells[Current_row, Current_col].Style.Font.Color;

     ATableView.Hyperlinks.Delete(ATableView.Hyperlinks[0]);

     DrawCellOnlyFont(SS,
                      0,
                      Current_row+1,
                      Current_col+1,
                      FontName,
                      FontSize,
                      FontStyle,
                      FontColor);
    end;
    SS.Sheets[Worksheet].EndUpdate;
  end;
end;

procedure CopyHyperlinks_from_Temp(SS:TdxSpreadSheet; Worksheet:integer);
var
 ATableView: TdxSpreadSheetTableView;
 temp_ATableView: TdxSpreadSheetTableView;
 temp_APictureContainer: TdxSpreadSheetPictureContainer;
 i:integer;
begin
  if (temp_SS<>nil) then
  begin
    temp_ATableView := TdxSpreadSheetTableView(temp_SS.Sheets[Worksheet]);
    if (temp_ATableView.Hyperlinks.Count>0) then
    begin
      ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);

      SS.Sheets[Worksheet].BeginUpdate;
      for i:=0 to temp_ATableView.Hyperlinks.Count-1 do
      begin
        temp_APictureContainer := ATableView.Containers.Add(TdxSpreadSheetPictureContainer) as TdxSpreadSheetPictureContainer;
        temp_APictureContainer.Hyperlink := ATableView.Hyperlinks.Add(temp_ATableView.Hyperlinks[i].Area);
        temp_APictureContainer.Hyperlink.Value := temp_ATableView.Hyperlinks[i].Value;
        temp_APictureContainer.Hyperlink.ScreenTip:=temp_ATableView.Hyperlinks[i].ScreenTip;
      end;
      SS.Sheets[Worksheet].EndUpdate;

      FreeAndNil(temp_APictureContainer);
      FreeAndNil(temp_ATableView);
      FreeAndNil(temp_SS);
    end;
  end;
end;

procedure RecalcFormulsWorksheet(SS:TdxSpreadSheet);
begin
  SS.FormulaController.Calculate;
  SS.Update;
end;


procedure ExelSaveToFile(SS:TdxSpreadSheet; FileName:string);
begin
  SS.SaveToFile(FileName);
end;

procedure ClearCells_(SS:TdxSpreadSheet; Worksheet:integer; col_left:integer; col_right:integer; row_up:integer; row_down:integer);
var
 ATableView: TdxSpreadSheetTableView;
 ARect: TRect;
begin
  ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
  ARect := Rect(col_left-1, row_up-1, col_right-1, row_down-1);

  ATableView.BeginUpdate; // Оптимизация прорисовки
  try
    // Метод принимает TRect и флаги очистки
    ATableView.ClearCells(ARect);
  finally
    ATableView.EndUpdate;
  end;
end;

procedure SetFocusedCell_(SS:TdxSpreadSheet; Worksheet:integer; col:integer; row:integer);
var
 ATableView: TdxSpreadSheetTableView;
begin
  ATableView := TdxSpreadSheetTableView(SS.Sheets[Worksheet]);
  ATableView.Selection.FocusedColumn := col-1;
  ATableView.Selection.FocusedRow := row-1;
end;




end.
