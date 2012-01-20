object frmEntradas: TfrmEntradas
  Left = 138
  Top = 62
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Entradas'
  ClientHeight = 615
  ClientWidth = 778
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsMDIChild
  Icon.Data = {
    0000010001001010000001002000680400001600000028000000100000002000
    000001002000000000000000000000000000000000000000000000000000FFFF
    FF2FFFFFFFDFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
    FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFDFFFFFFF2FFFFF
    FFDF85B785FF0F6F0FFF167416FF1A761AFF1A761AFF187818FF177917FF137D
    13FF0D7F0DFF0A7E0AFF077C07FF027B02FF007000FF7FB07FFFFFFFFFDFFFFF
    FFFF118311FF1F8C1FFF2A912AFF2F932FFF2E942EFF2C962CFF299A29FF239E
    23FF1CA31CFF15A415FF0DA40DFF059F05FF019101FF006F00FFFFFFFFFFFFFF
    FFFF198D19FF2C962CFF379C37FF3D9F3DFF3C9F3CFF39A139FFA3D6A3FFFFFF
    FFFF24AF24FF1CB11CFF13B213FF0AAD0AFF049F04FF027902FFFFFFFFFFFFFF
    FFFF229122FF389C38FF43A243FF48A448FF45A545FF42A642FFFFFFFFFFFFFF
    FFFFFFFFFFFF21B521FF18B618FF0EB10EFF08A308FF057E05FFFFFFFFFFFFFF
    FFFF2C962CFF42A042FF4CA54CFF4FA74FFF4CA74CFF46A746FF40AA40FFFFFF
    FFFFFFFFFFFFFFFFFFFF1AB31AFF14AF14FF0FA30FFF0B800BFFFFFFFFFFFFFF
    FFFF359A35FF4BA54BFF52A852FF53A953FF4EA84EFF49A749FF41A841FF38AA
    38FFFFFFFFFFFFFFFFFFFFFFFFFF19AC19FF18A218FF128212FFFFFFFFFFFFFF
    FFFF3F9F3FFF53A953FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
    FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF1F9E1FFF188118FFFFFFFFFFFFFF
    FFFF45A245FF5AAC5AFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
    FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF259A25FF1D7F1DFFFFFFFFFFFFFF
    FFFF4FA74FFF63B163FF61AF61FF59AB59FF51A651FF48A248FF3F9F3FFF369C
    36FFFFFFFFFFFFFFFFFFFFFFFFFF269926FF2A972AFF217E21FFFFFFFFFFFFFF
    FFFF53A953FF6CB66CFF68B468FF5EAD5EFF54A854FF4CA34CFF429F42FFFFFF
    FFFFFFFFFFFFFFFFFFFF299729FF2B982BFF2D952DFF237E23FFFFFFFFFFFFFF
    FFFF5EAF5EFF7ABD7AFF70B870FF63B063FF5AAB5AFF52A652FFFFFFFFFFFFFF
    FFFFFFFFFFFF339933FF309930FF309830FF2F942FFF237D23FFFFFFFFFFFFFF
    FFFF6BB56BFF8DC68DFF80C080FF6FB76FFF67B267FF60AE60FFB4D9B4FFFFFF
    FFFF4CA54CFF49A449FF41A141FF3A9D3AFF309530FF1E7A1EFFFFFFFFFFFFFF
    FFFF77BB77FF9DCF9DFF8CC68CFF79BC79FF70B870FF69B469FF65B265FF62B0
    62FF5DAE5DFF56AB56FF4EA74EFF41A141FF2F942FFF197719FFFFFFFFFFFFFF
    FFDFB1D8B1FF76BB76FF67B367FF5BAD5BFF54A954FF4FA74FFF4AA44AFF4BA5
    4BFF46A346FF3FA03FFF3B9E3BFF319831FF238C23FF8ABB8AFFFFFFFFDFFFFF
    FF2FFFFFFFDFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
    FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFDFFFFFFF2F0000
    0000000000000000000000000000000000000000000000000000000000000000
    000000000000000000000000000000000000000000000000000000000000}
  OldCreateOrder = False
  Position = poDefault
  Visible = True
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object gbButtons: TGroupBox
    Left = 3
    Top = 1
    Width = 772
    Height = 146
    TabOrder = 0
    object Label1: TLabel
      Left = 13
      Top = 18
      Width = 59
      Height = 13
      Caption = 'Pedimiento :'
    end
    object Label2: TLabel
      Left = 213
      Top = 19
      Width = 87
      Height = 13
      Caption = 'Clave Pedimento :'
    end
    object Label3: TLabel
      Left = 480
      Top = 19
      Width = 36
      Height = 13
      Caption = 'Fecha :'
    end
    object Label4: TLabel
      Left = 12
      Top = 46
      Width = 61
      Height = 13
      Caption = 'Pais Origen :'
    end
    object Label9: TLabel
      Left = 29
      Top = 84
      Width = 44
      Height = 13
      Caption = 'Factura :'
    end
    object OrdenCompra: TLabel
      Left = 223
      Top = 84
      Width = 77
      Height = 13
      Caption = 'Orden Compra :'
    end
    object lblProvedor: TLabel
      Left = 464
      Top = 84
      Width = 51
      Height = 13
      Caption = 'Provedor :'
    end
    object IVA: TLabel
      Left = 671
      Top = 82
      Width = 24
      Height = 13
      Caption = 'IVA :'
    end
    object lblId: TLabel
      Left = 8
      Top = 102
      Width = 8
      Height = 13
      Caption = 'id'
      Visible = False
    end
    object lblAnio: TLabel
      Left = 281
      Top = 118
      Width = 31
      Height = 13
      Caption = 'lblAnio'
    end
    object Label5: TLabel
      Left = 232
      Top = 46
      Width = 68
      Height = 13
      Caption = 'Tipo Entrada :'
    end
    object Label6: TLabel
      Left = 451
      Top = 39
      Width = 64
      Height = 26
      Caption = 'Tipo Importacion :'
      WordWrap = True
    end
    object Label11: TLabel
      Left = 653
      Top = 39
      Width = 42
      Height = 26
      Caption = 'Tipo Cambio :'
      WordWrap = True
    end
    object Nuevo: TButton
      Left = 14
      Top = 113
      Width = 45
      Height = 22
      Hint = 'Nueva Factura'
      Caption = 'Nuevo'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 11
      OnClick = NuevoClick
    end
    object Editar: TButton
      Left = 62
      Top = 113
      Width = 45
      Height = 22
      Hint = 'Editar Factura'
      Caption = 'Editar'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 12
      OnClick = EditarClick
    end
    object Borrar: TButton
      Left = 110
      Top = 113
      Width = 45
      Height = 22
      Hint = 'Borrar Factura'
      Caption = 'Borrar'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 13
      OnClick = BorrarClick
    end
    object Buscar: TButton
      Left = 157
      Top = 113
      Width = 45
      Height = 22
      Hint = 'Buscar Factura'
      Caption = 'Buscar'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 14
    end
    object Imprimir: TButton
      Left = 205
      Top = 113
      Width = 45
      Height = 22
      Hint = 'Aceptar Operacion'
      Caption = 'Imprimir'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 15
    end
    object Panel1: TPanel
      Left = 646
      Top = 110
      Width = 114
      Height = 27
      TabOrder = 18
      object Primero: TButton
        Left = 8
        Top = 4
        Width = 25
        Height = 19
        Hint = 'Primero'
        Caption = '| <'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = PrimeroClick
      end
      object Anterior: TButton
        Left = 32
        Top = 4
        Width = 25
        Height = 19
        Hint = 'Previo'
        Caption = '<'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = PrimeroClick
      end
      object Ultimo: TButton
        Left = 80
        Top = 4
        Width = 25
        Height = 19
        Hint = 'Ultimo'
        Caption = '> |'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = PrimeroClick
      end
      object Siguiente: TButton
        Left = 56
        Top = 4
        Width = 25
        Height = 19
        Hint = 'Siguiente'
        Caption = '>'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = PrimeroClick
      end
    end
    object btnAceptar: TButton
      Left = 526
      Top = 113
      Width = 45
      Height = 22
      Hint = 'Aceptar Operacion'
      Caption = 'Grabar'
      Enabled = False
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 16
      OnClick = btnAceptarClick
    end
    object btnCancelar: TButton
      Left = 575
      Top = 113
      Width = 48
      Height = 22
      Hint = 'Cancelar Operacion'
      Caption = 'Cancelar'
      Enabled = False
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 17
      OnClick = btnCancelarClick
    end
    object txtPedimento: TEdit
      Left = 76
      Top = 16
      Width = 130
      Height = 21
      ReadOnly = True
      TabOrder = 0
      OnKeyDown = SendTab
    end
    object txtClavePedimento: TEdit
      Left = 303
      Top = 16
      Width = 130
      Height = 21
      ReadOnly = True
      TabOrder = 1
      OnKeyDown = SendTab
    end
    object txtFactura: TEdit
      Left = 76
      Top = 80
      Width = 130
      Height = 21
      ReadOnly = True
      TabOrder = 7
      OnKeyDown = SendTab
    end
    object txtOrdenCompra: TEdit
      Left = 303
      Top = 80
      Width = 130
      Height = 21
      ReadOnly = True
      TabOrder = 8
      OnKeyDown = SendTab
    end
    object deFecha: TDateEditor
      Left = 518
      Top = 16
      Width = 130
      Height = 20
      Alignment = taLeftJustify
      BorderStyle = bsSingle
      Enabled = False
      Margin = 0
      ParentColor = False
      TabOrder = 2
      TabStop = True
      VerticalAlignment = vaMiddle
      OnKeyDown = SendTab
      AutoSelect = False
      ReadOnly = False
      Text = '9/9/2007'
      Style = dsDropDown
      Date = 39334
    end
    object ddlPais: TComboBox
      Left = 76
      Top = 44
      Width = 130
      Height = 21
      Style = csDropDownList
      Enabled = False
      ItemHeight = 13
      TabOrder = 4
      OnKeyDown = SendTab
    end
    object ddlProvedor: TComboBox
      Left = 518
      Top = 80
      Width = 130
      Height = 21
      Style = csDropDownList
      Enabled = False
      ItemHeight = 13
      TabOrder = 9
      OnKeyDown = SendTab
    end
    object txtIVA: TEdit
      Left = 697
      Top = 78
      Width = 63
      Height = 21
      MaxLength = 2
      ReadOnly = True
      TabOrder = 10
      OnKeyDown = SendTab
      OnKeyPress = txtIVAKeyPress
    end
    object ddlNacional: TComboBox
      Left = 303
      Top = 44
      Width = 130
      Height = 21
      Style = csDropDownList
      Enabled = False
      ItemHeight = 13
      TabOrder = 5
      OnChange = ddlNacionalChange
      OnKeyDown = SendTab
      Items.Strings = (
        'Importado'
        'Nacional')
    end
    object ddlTipoImp: TComboBox
      Left = 518
      Top = 44
      Width = 130
      Height = 21
      Style = csDropDownList
      Enabled = False
      ItemHeight = 13
      TabOrder = 6
      OnKeyDown = SendTab
      Items.Strings = (
        'Importacion Temporal'
        'Importacion Definitiva')
    end
    object chkDlls: TCheckBox
      Left = 698
      Top = 15
      Width = 63
      Height = 17
      Caption = 'Dolares'
      Enabled = False
      TabOrder = 3
      OnKeyDown = SendTab
    end
    object txtTipo: TEdit
      Left = 698
      Top = 44
      Width = 63
      Height = 21
      ReadOnly = True
      TabOrder = 19
      OnKeyDown = SendTab
      OnKeyPress = txtTasaGeneralKeyPress
    end
  end
  object GroupBox2: TGroupBox
    Left = 4
    Top = 147
    Width = 771
    Height = 465
    TabOrder = 1
    object lblLongitud: TLabel
      Left = 479
      Top = 73
      Width = 48
      Height = 13
      Caption = 'Longitud :'
      Visible = False
    end
    object lblDiametro: TLabel
      Left = 298
      Top = 73
      Width = 50
      Height = 13
      Caption = 'Diametro :'
      Visible = False
    end
    object Label7: TLabel
      Left = 17
      Top = 44
      Width = 56
      Height = 13
      Caption = 'Material ID:'
    end
    object Label8: TLabel
      Left = 23
      Top = 73
      Width = 50
      Height = 13
      Caption = 'Cantidad :'
    end
    object Label10: TLabel
      Left = 162
      Top = 73
      Width = 35
      Height = 13
      Caption = 'Costo :'
    end
    object Label13: TLabel
      Left = 587
      Top = 392
      Width = 46
      Height = 13
      Caption = 'Sub-Total'
    end
    object Label14: TLabel
      Left = 608
      Top = 416
      Width = 25
      Height = 13
      Caption = 'I.V.A'
    end
    object Label15: TLabel
      Left = 604
      Top = 440
      Width = 32
      Height = 13
      Caption = 'Total '
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lblMaterialDesc: TLabel
      Left = 8
      Top = 392
      Width = 71
      Height = 13
      Caption = 'lblMaterialDesc'
    end
    object Label12: TLabel
      Left = 5
      Top = 20
      Width = 68
      Height = 13
      Caption = 'Proveedor ID:'
    end
    object lblPies: TLabel
      Left = 604
      Top = 90
      Width = 16
      Height = 11
      Caption = 'pies'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -9
      Font.Name = 'Tahoma'
      Font.Style = [fsItalic]
      ParentFont = False
      Visible = False
    end
    object lblPulgadas: TLabel
      Left = 413
      Top = 90
      Width = 36
      Height = 11
      Caption = 'pulgadas'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -9
      Font.Name = 'Tahoma'
      Font.Style = [fsItalic]
      ParentFont = False
      Visible = False
    end
    object txtLongitud: TEdit
      Left = 532
      Top = 69
      Width = 85
      Height = 21
      Enabled = False
      TabOrder = 5
      Visible = False
      OnKeyDown = SendTab
      OnKeyPress = txtTasaGeneralKeyPress
    end
    object txtDiametro: TEdit
      Left = 351
      Top = 69
      Width = 85
      Height = 21
      Enabled = False
      TabOrder = 4
      Visible = False
      OnKeyDown = SendTab
      OnKeyPress = txtTasaGeneralKeyPress
    end
    object btnClear: TButton
      Left = 423
      Top = 69
      Width = 39
      Height = 20
      Hint = 'Aceptar Operacion'
      Caption = 'Limpiar'
      Enabled = False
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 8
      OnClick = btnClearClick
    end
    object btnDelete: TButton
      Left = 377
      Top = 69
      Width = 42
      Height = 20
      Hint = 'Aceptar Operacion'
      Caption = 'Borrar'
      Enabled = False
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 7
      OnClick = btnDeleteClick
    end
    object btnAdd: TButton
      Left = 324
      Top = 69
      Width = 49
      Height = 20
      Hint = 'Aceptar Operacion'
      Caption = 'Agregar'
      Enabled = False
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
      OnClick = btnAddClick
    end
    object gvEntradas: TGridView
      Left = 8
      Top = 102
      Width = 754
      Height = 272
      BorderStyle = bsSingle
      Enabled = False
      GridStyle = gsReport
      GridLinesStyle = lsNormal
      HeaderSize = 18
      HeaderStyle = hsAuto
      HideScrollBar = True
      InputSize = 16
      Options = [goAlwaysShowSelection, goHeader, goHighlightTextSelection, goSelectFullRow]
      RowSize = 16
      ParentColor = False
      SelectionMoveDirection = mdDown
      SlideSize = 80
      TabOrder = 9
      TabStop = True
      WantReturns = False
      OnAfterEdit = gvEntradasAfterEdit
      object TTextualColumn
        Alignment = taLeftJustify
        Color = clWindow
        Cursor = crDefault
        DrawingOptions = doNormal
        DefaultWidth = 80
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Footer.Color = clWindow
        Footer.FormulaKind = fkNone
        Header.Color = clBtnFace
        Header.Caption = 'Id'
        Header.DisplayMode = dmTextOnly
        Options = [coCanClick, coCanSort, coEditorAutoSelect, coPublicUsing]
        Position = 0
        SlideBounds.Height = 0
        SlideBounds.Left = 0
        SlideBounds.Top = 0
        SlideBounds.Width = 0
        Sorted = False
        SortType = stAlphabetic
        VerticalAlignment = vaMiddle
        Visible = False
        Width = 80
        WrapKind = wkEllipsis
        AutoExecute = False
      end
      object TTextualColumn
        Alignment = taLeftJustify
        Color = clWindow
        Cursor = crDefault
        DrawingOptions = doNormal
        DefaultWidth = 95
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Footer.Color = clWindow
        Footer.FormulaKind = fkNone
        Header.Color = clBtnFace
        Header.Caption = 'ID'
        Header.DisplayMode = dmTextOnly
        Options = [coCanClick, coCanSort, coEditorAutoSelect, coPublicUsing]
        Position = 1
        SlideBounds.Height = 0
        SlideBounds.Left = 0
        SlideBounds.Top = 0
        SlideBounds.Width = 0
        Sorted = False
        SortType = stAlphabetic
        VerticalAlignment = vaMiddle
        Visible = True
        Width = 95
        WrapKind = wkEllipsis
        AutoExecute = False
      end
      object TTextualColumn
        Alignment = taLeftJustify
        Color = clWindow
        Cursor = crDefault
        DrawingOptions = doNormal
        DefaultWidth = 95
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Footer.Color = clWindow
        Footer.FormulaKind = fkNone
        Header.Color = clBtnFace
        Header.Caption = 'Proveedor ID'
        Header.DisplayMode = dmTextOnly
        Options = [coCanClick, coCanSort, coEditorAutoSelect, coPublicUsing]
        Position = 2
        SlideBounds.Height = 0
        SlideBounds.Left = 0
        SlideBounds.Top = 0
        SlideBounds.Width = 0
        Sorted = False
        SortType = stAlphabetic
        VerticalAlignment = vaMiddle
        Visible = True
        Width = 95
        WrapKind = wkEllipsis
        AutoExecute = False
      end
      object TTextualColumn
        Alignment = taLeftJustify
        Color = clWindow
        Cursor = crDefault
        DrawingOptions = doNormal
        DefaultWidth = 350
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Footer.Color = clWindow
        Footer.FormulaKind = fkNone
        Header.Color = clBtnFace
        Header.Caption = 'Material'
        Header.DisplayMode = dmTextOnly
        Options = [coCanClick, coCanSort, coEditorAutoSelect, coPublicUsing]
        Position = 3
        SlideBounds.Height = 0
        SlideBounds.Left = 0
        SlideBounds.Top = 0
        SlideBounds.Width = 0
        Sorted = True
        SortType = stAlphabetic
        VerticalAlignment = vaMiddle
        Visible = True
        Width = 350
        WrapKind = wkEllipsis
        AutoExecute = False
      end
      object TTextualColumn
        Alignment = taRightJustify
        Color = clWindow
        Cursor = crDefault
        DrawingOptions = doNormal
        DefaultWidth = 65
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Footer.Color = clWindow
        Footer.FormulaKind = fkNone
        Header.Color = clBtnFace
        Header.Caption = 'Cantidad'
        Header.DisplayMode = dmTextOnly
        Options = [coCanClick, coCanSort, coEditing, coEditorAutoSelect, coPublicUsing]
        Position = 4
        SlideBounds.Height = 0
        SlideBounds.Left = 0
        SlideBounds.Top = 0
        SlideBounds.Width = 0
        Sorted = False
        SortType = stAlphabetic
        VerticalAlignment = vaMiddle
        Visible = True
        Width = 65
        WrapKind = wkEllipsis
        AutoExecute = False
      end
      object TTextualColumn
        Alignment = taRightJustify
        Color = clWindow
        Cursor = crDefault
        DrawingOptions = doNormal
        DefaultWidth = 65
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Footer.Color = clWindow
        Footer.FormulaKind = fkNone
        Header.Color = clBtnFace
        Header.Caption = 'Costo'
        Header.DisplayMode = dmTextOnly
        Options = [coCanClick, coCanSort, coEditing, coEditorAutoSelect, coPublicUsing]
        Position = 5
        SlideBounds.Height = 0
        SlideBounds.Left = 0
        SlideBounds.Top = 0
        SlideBounds.Width = 0
        Sorted = False
        SortType = stAlphabetic
        VerticalAlignment = vaMiddle
        Visible = True
        Width = 65
        WrapKind = wkEllipsis
        AutoExecute = False
      end
      object TTextualColumn
        Alignment = taRightJustify
        Color = clWindow
        Cursor = crDefault
        DrawingOptions = doNormal
        DefaultWidth = 65
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Footer.Color = clWindow
        Footer.FormulaKind = fkNone
        Header.Color = clBtnFace
        Header.Caption = 'Total'
        Header.DisplayMode = dmTextOnly
        Options = [coCanClick, coEditorAutoSelect, coPublicUsing]
        Position = 6
        SlideBounds.Height = 0
        SlideBounds.Left = 0
        SlideBounds.Top = 0
        SlideBounds.Width = 0
        Sorted = False
        SortType = stAlphabetic
        VerticalAlignment = vaMiddle
        Visible = True
        Width = 65
        WrapKind = wkEllipsis
        AutoExecute = False
      end
      object TTextualColumn
        Alignment = taLeftJustify
        Color = clWindow
        Cursor = crDefault
        DrawingOptions = doNormal
        DefaultWidth = 80
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Footer.Color = clWindow
        Footer.FormulaKind = fkNone
        Header.Color = clBtnFace
        Header.Caption = 'MaterialID'
        Header.DisplayMode = dmTextOnly
        Options = [coCanClick, coCanSort, coEditorAutoSelect, coPublicUsing]
        Position = 7
        SlideBounds.Height = 0
        SlideBounds.Left = 0
        SlideBounds.Top = 0
        SlideBounds.Width = 0
        Sorted = False
        SortType = stAlphabetic
        VerticalAlignment = vaMiddle
        Visible = False
        Width = 80
        WrapKind = wkEllipsis
        AutoExecute = False
      end
      object TTextualColumn
        Alignment = taLeftJustify
        Color = clWindow
        Cursor = crDefault
        DrawingOptions = doNormal
        DefaultWidth = 80
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Footer.Color = clWindow
        Footer.FormulaKind = fkNone
        Header.Color = clBtnFace
        Header.Caption = 'Restante'
        Header.DisplayMode = dmTextOnly
        Options = [coCanClick, coCanSort, coEditorAutoSelect, coPublicUsing]
        Position = 8
        SlideBounds.Height = 0
        SlideBounds.Left = 0
        SlideBounds.Top = 0
        SlideBounds.Width = 0
        Sorted = False
        SortType = stAlphabetic
        VerticalAlignment = vaMiddle
        Visible = False
        Width = 80
        WrapKind = wkEllipsis
        AutoExecute = False
      end
      object TTextualColumn
        Alignment = taLeftJustify
        Color = clWindow
        Cursor = crDefault
        DrawingOptions = doNormal
        DefaultWidth = 80
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Footer.Color = clWindow
        Footer.FormulaKind = fkNone
        Header.Color = clBtnFace
        Header.Caption = 'Actual'
        Header.DisplayMode = dmTextOnly
        Options = [coCanClick, coCanSort, coEditorAutoSelect, coPublicUsing]
        Position = 9
        SlideBounds.Height = 0
        SlideBounds.Left = 0
        SlideBounds.Top = 0
        SlideBounds.Width = 0
        Sorted = False
        SortType = stAlphabetic
        VerticalAlignment = vaMiddle
        Visible = False
        Width = 80
        WrapKind = wkEllipsis
        AutoExecute = False
      end
    end
    object ddlMaterial: TComboBox
      Left = 75
      Top = 43
      Width = 177
      Height = 21
      Style = csDropDownList
      Enabled = False
      ItemHeight = 13
      TabOrder = 13
      Visible = False
      OnKeyDown = SendTab
    end
    object txtCantidad: TEdit
      Left = 75
      Top = 69
      Width = 85
      Height = 21
      ReadOnly = True
      TabOrder = 2
      OnKeyDown = SendTab
      OnKeyPress = txtTasaGeneralKeyPress
    end
    object txtCosto: TEdit
      Left = 199
      Top = 69
      Width = 85
      Height = 21
      ReadOnly = True
      TabOrder = 3
      OnKeyDown = SendTab
      OnKeyPress = txtTasaGeneralKeyPress
    end
    object txtSubtotal: TEdit
      Left = 639
      Top = 388
      Width = 121
      Height = 21
      ReadOnly = True
      TabOrder = 10
    end
    object txtTIVA: TEdit
      Left = 639
      Top = 412
      Width = 121
      Height = 21
      ReadOnly = True
      TabOrder = 11
    end
    object txtTotal: TEdit
      Left = 639
      Top = 436
      Width = 121
      Height = 21
      ReadOnly = True
      TabOrder = 12
    end
    object txtID: TEdit
      Left = 75
      Top = 43
      Width = 689
      Height = 21
      MaxLength = 50
      ReadOnly = True
      TabOrder = 1
      OnExit = txtIDExit
      OnKeyDown = txtIDKeyDown
    end
    object btnOrden: TButton
      Left = 745
      Top = 45
      Width = 17
      Height = 17
      BiDiMode = bdLeftToRight
      Caption = 'q'
      Enabled = False
      Font.Charset = SYMBOL_CHARSET
      Font.Color = clWindowText
      Font.Height = -8
      Font.Name = 'Wingdings 3'
      Font.Style = []
      ParentBiDiMode = False
      ParentFont = False
      TabOrder = 15
      OnClick = btnOrdenClick
    end
    object txtProvId: TEdit
      Left = 75
      Top = 16
      Width = 169
      Height = 21
      ReadOnly = True
      TabOrder = 0
      OnKeyDown = txtProvIdKeyDown
    end
    object gbMateriales: TGroupBox
      Left = 39
      Top = 136
      Width = 690
      Height = 315
      TabOrder = 14
      Visible = False
      object tvMateriales: TTreeView
        Left = 0
        Top = 1
        Width = 688
        Height = 312
        Indent = 19
        TabOrder = 0
        OnDblClick = tvMaterialesDblClick
      end
    end
  end
end
