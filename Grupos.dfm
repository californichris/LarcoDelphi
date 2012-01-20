object frmGrupos: TfrmGrupos
  Left = 431
  Top = 140
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Grupos'
  ClientHeight = 324
  ClientWidth = 227
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsMDIChild
  OldCreateOrder = False
  Position = poDefault
  Visible = True
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object gvGrupos: TGridView
    Left = 8
    Top = 74
    Width = 209
    Height = 241
    BorderStyle = bsSingle
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    GridStyle = gsReport
    GridLinesStyle = lsNormal
    HeaderSize = 18
    HeaderStyle = hsAuto
    HideScrollBar = True
    InputSize = 16
    Options = [goHeader, goHighlightTextSelection, goSelectFullRow]
    RowSize = 16
    ParentColor = False
    ParentFont = False
    PopupMenu = PopupMenu1
    SelectionMoveDirection = mdDown
    SlideSize = 80
    TabOrder = 0
    TabStop = True
    WantReturns = False
    object TTextualColumn
      Alignment = taLeftJustify
      Color = clWindow
      Cursor = crDefault
      DrawingOptions = doNormal
      DefaultWidth = 30
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
      Width = 30
      WrapKind = wkEllipsis
      AutoExecute = False
    end
    object TTextualColumn
      Alignment = taLeftJustify
      Color = clWindow
      Cursor = crDefault
      DrawingOptions = doNormal
      DefaultWidth = 200
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      Footer.Color = clWindow
      Footer.FormulaKind = fkNone
      Header.Color = clBtnFace
      Header.Caption = 'Nombre'
      Header.DisplayMode = dmTextOnly
      Options = [coCanClick, coCanSort, coEditorAutoSelect, coPublicUsing]
      Position = 1
      SlideBounds.Height = 0
      SlideBounds.Left = 0
      SlideBounds.Top = 0
      SlideBounds.Width = 0
      Sorted = True
      SortType = stAlphabetic
      VerticalAlignment = vaMiddle
      Visible = True
      Width = 200
      WrapKind = wkEllipsis
      AutoExecute = False
    end
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 0
    Width = 209
    Height = 65
    TabOrder = 1
    object Label1: TLabel
      Left = 11
      Top = 16
      Width = 47
      Height = 13
      Alignment = taRightJustify
      Caption = 'Nombre : '
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object txtNombre: TEdit
      Left = 64
      Top = 12
      Width = 137
      Height = 21
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      MaxLength = 50
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
    end
    object Button1: TButton
      Left = 108
      Top = 38
      Width = 41
      Height = 21
      Caption = 'Grabar'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      OnClick = Button1Click
    end
    object btnCancelar: TButton
      Left = 152
      Top = 38
      Width = 48
      Height = 21
      Caption = 'Cancelar'
      Enabled = False
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      OnClick = btnCancelarClick
    end
  end
  object PopupMenu1: TPopupMenu
    Left = 16
    Top = 96
    object Borrar1: TMenuItem
      Caption = 'Borrar'
      OnClick = Borrar1Click
    end
    object Editar1: TMenuItem
      Caption = 'Modificar'
      OnClick = Editar1Click
    end
  end
end
