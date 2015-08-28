VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hammer launcher"
   ClientHeight    =   6840
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6075
   Icon            =   "hammer launcher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0.93
   ScaleMode       =   0  'User
   ScaleWidth      =   1
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame ChoiceFrame 
      BorderStyle     =   0  'None
      Height          =   5055
      Index           =   3
      Left            =   240
      TabIndex        =   21
      Top             =   1080
      Width           =   5655
      Begin VB.CommandButton AddWad 
         Caption         =   "Add &WAD"
         Height          =   375
         Left            =   3120
         TabIndex        =   35
         ToolTipText     =   "Add a new WAD the list."
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton WadRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         TabIndex        =   36
         ToolTipText     =   "Remove the selected wad from the list."
         Top             =   120
         Width           =   1215
      End
      Begin VB.ListBox WadList 
         Height          =   4350
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   34
         ToolTipText     =   "List of WAD's that will be loaded in to hammer."
         Top             =   600
         Width           =   5655
      End
      Begin VB.Label Label2 
         Caption         =   " Hammer's WAD file list."
         Height          =   255
         Left            =   0
         TabIndex        =   37
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame ChoiceFrame 
      BorderStyle     =   0  'None
      Height          =   5055
      Index           =   4
      Left            =   240
      TabIndex        =   58
      Top             =   1080
      Width           =   5655
      Begin VB.Frame Frame10 
         Caption         =   "Bars"
         Height          =   1815
         Left            =   120
         TabIndex        =   84
         Top             =   3120
         Width           =   5415
         Begin VB.CheckBox Check11 
            Caption         =   "Check1"
            Height          =   195
            Left            =   2760
            TabIndex        =   95
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Check1"
            Height          =   195
            Left            =   2760
            TabIndex        =   94
            Top             =   480
            Width           =   1935
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Check1"
            Height          =   195
            Left            =   2760
            TabIndex        =   93
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Check1"
            Height          =   195
            Left            =   2760
            TabIndex        =   92
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Check1"
            Height          =   195
            Left            =   2760
            TabIndex        =   91
            Top             =   960
            Width           =   1935
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Check1"
            Height          =   195
            Left            =   240
            TabIndex        =   90
            Top             =   1440
            Width           =   1935
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Check1"
            Height          =   195
            Left            =   240
            TabIndex        =   89
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Check1"
            Height          =   195
            Left            =   240
            TabIndex        =   88
            Top             =   960
            Width           =   1935
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Check1"
            Height          =   195
            Left            =   240
            TabIndex        =   87
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Check1"
            Height          =   195
            Left            =   240
            TabIndex        =   86
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   195
            Left            =   240
            TabIndex        =   85
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Window Setup"
         Height          =   2895
         Left            =   120
         TabIndex        =   69
         Top             =   120
         Width           =   5415
         Begin VB.CheckBox IndependentWindows 
            Caption         =   "Use &independent window configurations"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            ToolTipText     =   "Use individual windows instead of a set ot 2x2."
            Top             =   240
            Width           =   3255
         End
         Begin VB.CheckBox LoadDefaultPositions 
            Caption         =   "Load default window &positions with maps"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   70
            ToolTipText     =   "Load the custom window setup When a map gets loaded."
            Top             =   600
            Value           =   1  'Checked
            Width           =   3255
         End
         Begin VB.ComboBox view1 
            Height          =   315
            ItemData        =   "hammer launcher.frx":2AFA
            Left            =   120
            List            =   "hammer launcher.frx":2B10
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   960
            Width           =   1215
         End
         Begin VB.ComboBox view2 
            Height          =   315
            ItemData        =   "hammer launcher.frx":2B58
            Left            =   2760
            List            =   "hammer launcher.frx":2B6E
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   960
            Width           =   1215
         End
         Begin VB.ComboBox view3 
            Height          =   315
            ItemData        =   "hammer launcher.frx":2BB6
            Left            =   120
            List            =   "hammer launcher.frx":2BCC
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   1920
            Width           =   1215
         End
         Begin VB.ComboBox view4 
            Height          =   315
            ItemData        =   "hammer launcher.frx":2C14
            Left            =   2760
            List            =   "hammer launcher.frx":2C2A
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   1920
            Width           =   1215
         End
         Begin VB.PictureBox ImgTex 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   4440
            Picture         =   "hammer launcher.frx":2C72
            ScaleHeight     =   255
            ScaleWidth      =   735
            TabIndex        =   79
            Top             =   600
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.PictureBox ImgSolid 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   4440
            Picture         =   "hammer launcher.frx":37DC
            ScaleHeight     =   255
            ScaleWidth      =   735
            TabIndex        =   78
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.PictureBox Img2D 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3600
            Picture         =   "hammer launcher.frx":3950
            ScaleHeight     =   255
            ScaleWidth      =   735
            TabIndex        =   74
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.PictureBox ImgWire 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3600
            Picture         =   "hammer launcher.frx":3FBF
            ScaleHeight     =   255
            ScaleWidth      =   735
            TabIndex        =   73
            Top             =   600
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.PictureBox Picture1 
            Height          =   855
            Left            =   120
            ScaleHeight     =   795
            ScaleWidth      =   2475
            TabIndex        =   80
            Top             =   960
            Width           =   2535
         End
         Begin VB.PictureBox Picture3 
            Height          =   855
            Left            =   120
            ScaleHeight     =   795
            ScaleWidth      =   2475
            TabIndex        =   81
            Top             =   1920
            Width           =   2535
         End
         Begin VB.PictureBox Picture4 
            Height          =   855
            Left            =   2760
            ScaleHeight     =   795
            ScaleWidth      =   2475
            TabIndex        =   82
            Top             =   1920
            Width           =   2535
         End
         Begin VB.PictureBox Picture2 
            Height          =   855
            Left            =   2760
            ScaleHeight     =   795
            ScaleWidth      =   2475
            TabIndex        =   83
            Top             =   960
            Width           =   2535
         End
      End
   End
   Begin VB.Frame ChoiceFrame 
      BorderStyle     =   0  'None
      Height          =   4935
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   5655
      Begin VB.Frame Frame6 
         Caption         =   "Preformance"
         Height          =   1335
         Left            =   120
         TabIndex        =   40
         Top             =   120
         Width           =   5415
         Begin VB.CheckBox FilterTextures 
            Caption         =   "F&ilter textures"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            ToolTipText     =   "Removes pixilation of texturesl."
            Top             =   240
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox AnimateModels 
            Caption         =   "Animate &models"
            Height          =   255
            Left            =   2760
            TabIndex        =   43
            ToolTipText     =   "Animate the entity models."
            Top             =   240
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin MSComctlLib.Slider BackPlaneScroll 
            Height          =   255
            Left            =   1800
            TabIndex        =   54
            ToolTipText     =   "Distance in units that hammer will draw."
            Top             =   600
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   393216
            Max             =   16384
            SelStart        =   10000
            TickStyle       =   3
            Value           =   10000
         End
         Begin MSComctlLib.Slider ModelDistanceScroll 
            Height          =   255
            Left            =   1800
            TabIndex        =   55
            ToolTipText     =   "Distance in units that hammer will draw entitys as models."
            Top             =   960
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   393216
            Max             =   16384
            SelStart        =   2000
            TickStyle       =   3
            Value           =   2000
         End
         Begin VB.Label Label8 
            Caption         =   "Model render distance:"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            ToolTipText     =   "how fare away will entitys show as models."
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label ModelDistanceTxt 
            Alignment       =   1  'Right Justify
            Caption         =   "2000"
            Height          =   255
            Left            =   4800
            TabIndex        =   46
            Top             =   960
            Width           =   495
         End
         Begin VB.Label BackPlaneTxt 
            Alignment       =   1  'Right Justify
            Caption         =   "10000"
            Height          =   255
            Left            =   4800
            TabIndex        =   45
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "Back clipping plane:"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Navigation"
         Height          =   1335
         Left            =   120
         TabIndex        =   38
         Top             =   1560
         Width           =   5415
         Begin VB.CheckBox UseMouseLook 
            Caption         =   "Use mouselook &navigation"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            ToolTipText     =   "Enable spector like camera controle by pressing z or holding space."
            Top             =   240
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox ReverseY 
            Caption         =   "Revers mouse &Y axis"
            Height          =   255
            Left            =   2760
            TabIndex        =   48
            ToolTipText     =   "Invrt up and down on the camera."
            Top             =   240
            Width           =   1815
         End
         Begin MSComctlLib.Slider ForwardSpeedMaxScroll 
            Height          =   255
            Left            =   1080
            TabIndex        =   56
            ToolTipText     =   "How fast will the spector camera move."
            Top             =   600
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   450
            _Version        =   393216
            Min             =   100
            Max             =   10000
            SelStart        =   950
            TickStyle       =   3
            Value           =   950
         End
         Begin ComctlLib.Slider TimeToMaxSpeedScroll 
            Height          =   255
            Left            =   1560
            TabIndex        =   57
            ToolTipText     =   "Time it take to accelerate full move speed."
            Top             =   960
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   450
            _Version        =   327682
            Max             =   10000
            SelStart        =   550
            TickStyle       =   3
            Value           =   550
         End
         Begin VB.Label TimeToMaxSpeedTxt 
            Alignment       =   1  'Right Justify
            Caption         =   "0.55 sec."
            Height          =   255
            Left            =   4440
            TabIndex        =   52
            Top             =   960
            Width           =   855
         End
         Begin VB.Label ForwardSpeedMaxTxt 
            Alignment       =   1  'Right Justify
            Caption         =   "950"
            Height          =   255
            Left            =   4800
            TabIndex        =   51
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label10 
            Caption         =   "Time to top speed:"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label9 
            Caption         =   "Move speed:"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "General"
         Height          =   615
         Left            =   120
         TabIndex        =   41
         Top             =   3000
         Width           =   5415
         Begin VB.CheckBox ReverseSelection 
            Caption         =   "Revers &selectio order."
            Height          =   195
            Left            =   2760
            TabIndex        =   53
            ToolTipText     =   "Some grafic cards will cause the brushes to be selected in the wrong order."
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label ClearColorLabel 
            Caption         =   "&Background color"
            Height          =   255
            Left            =   405
            TabIndex        =   60
            ToolTipText     =   "Color of the void."
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label ClearColor 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Height          =   210
            Left            =   120
            TabIndex        =   59
            ToolTipText     =   "Color of the void."
            Top             =   240
            Width           =   210
         End
      End
   End
   Begin VB.Frame ChoiceFrame 
      BorderStyle     =   0  'None
      Height          =   5055
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   5655
      Begin VB.Frame Frame1 
         Caption         =   "Options"
         Height          =   3015
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   5415
         Begin VB.CheckBox RotateConstrain 
            Caption         =   "Limit to 15 degree &rotations"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            ToolTipText     =   "Makes it posible to make 90° rotation with less chance of bad brush and leaks."
            Top             =   240
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox Scrollbars 
            Caption         =   "Display &scrollbars"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   "Use scrollbars to navigate instead of space+drag."
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox DrawVertices 
            Caption         =   "Draw &vertices"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "Draw brush vertices(cornor)."
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox WhiteOnBlack 
            Caption         =   "&White-on-Black color scheme"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            ToolTipText     =   "White grid on black background."
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox KeepCloneGroup 
            Caption         =   "Keep &group when clone-dragging"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "Copy's of groups are also grouped."
            Top             =   2160
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox StretchArches 
            Caption         =   "Stretch arches to &fit original bounding rectangle"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            ToolTipText     =   "Scales  arches to the size of the rectangle used to create it."
            Top             =   2640
            Value           =   1  'Checked
            Width           =   3735
         End
         Begin VB.CheckBox Usegroupcolors 
            Caption         =   "Use visgroup &colors for object lines"
            Height          =   375
            Left            =   2640
            TabIndex        =   12
            ToolTipText     =   "If brushes are in a group they will recive the same color."
            Top             =   240
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox Nudge 
            Caption         =   "Arrow keys &nudge selected object/vertex"
            Height          =   375
            Left            =   2640
            TabIndex        =   13
            ToolTipText     =   "Alowes you to move brushes/vertices 1 grid size with the arrow key's."
            Top             =   720
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox OrientPrimitives 
            Caption         =   "Reorient &primitives/prefabs on creation in the active 2D view."
            Height          =   375
            Left            =   2640
            TabIndex        =   14
            ToolTipText     =   "Object will have there top facing to the view it was created in."
            Top             =   1200
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.CheckBox AutoSelect 
            Caption         =   "Automatic &infinite selection in 2D windows (no ENTER)"
            Height          =   375
            Left            =   2640
            TabIndex        =   15
            ToolTipText     =   "Make selection work like in windows."
            Top             =   1680
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.CheckBox SelectByHandles 
            Caption         =   "Selection box selects by center &handles only"
            Height          =   375
            Left            =   2640
            TabIndex        =   16
            ToolTipText     =   "When on, brushes can only be selected at there center."
            Top             =   2160
            Width           =   2415
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Grid"
         Height          =   1695
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   5415
         Begin VB.ComboBox DefaultGrid 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1030
               SubFormatType   =   1
            EndProperty
            Height          =   315
            ItemData        =   "hammer launcher.frx":43A6
            Left            =   600
            List            =   "hammer launcher.frx":43BF
            TabIndex        =   18
            Text            =   "8"
            ToolTipText     =   "Grid lines will be drawn on every N'th world unit."
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox GridHighSpec 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1030
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   22
            Text            =   "1"
            Top             =   600
            Width           =   495
         End
         Begin VB.CheckBox Gridhigh64 
            Caption         =   "Highlight every &64 units"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            ToolTipText     =   "Changes the intensity of grid lines on every 64'th world unit."
            Top             =   960
            Width           =   2055
         End
         Begin VB.CheckBox Gridhigh1024 
            Caption         =   "Highlight every &1024 units"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            ToolTipText     =   "Changes the intensity of grid lines on every 1024'th world unit."
            Top             =   1320
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin MSComctlLib.Slider GridIntensityscroll 
            Height          =   255
            Left            =   2880
            TabIndex        =   26
            Tag             =   "grid contrast to the back ground."
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            _Version        =   393216
            Max             =   100
            SelStart        =   30
            TickStyle       =   3
            Value           =   30
         End
         Begin VB.CheckBox GridDots 
            Caption         =   "&Dotted grid"
            Height          =   255
            Left            =   2880
            TabIndex        =   28
            Tag             =   "Only show where the grid intersects."
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox HideSmallGrid 
            Caption         =   "Hide grid smaller then &4 pixels"
            Height          =   195
            Left            =   2880
            TabIndex        =   24
            ToolTipText     =   "Gives a mutch cleares view."
            Top             =   960
            Value           =   1  'Checked
            Width           =   2415
         End
         Begin VB.Label Label1 
            Caption         =   "Highlight every"
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label GridIntensityTxt 
            Caption         =   "Intensity: 30%"
            Height          =   255
            Left            =   1800
            TabIndex        =   27
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Size:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   300
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "grid line."
            Height          =   255
            Left            =   1920
            TabIndex        =   23
            Top             =   600
            Width           =   735
         End
      End
   End
   Begin VB.Frame ChoiceFrame 
      BorderStyle     =   0  'None
      Height          =   5055
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   5655
      Begin VB.Frame Frame3 
         Caption         =   "Undo"
         Height          =   975
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   5415
         Begin MSComctlLib.Slider UndoLevelsScroll 
            Height          =   255
            Left            =   960
            TabIndex        =   32
            ToolTipText     =   "Number of undo's that you can preform."
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   450
            _Version        =   393216
            Max             =   999
            SelStart        =   50
            TickStyle       =   3
            Value           =   50
         End
         Begin VB.CheckBox UndoMemoryWarning 
            Caption         =   "Show &warning when low on  memory"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            ToolTipText     =   "Shows a warning if the system gets low on memory."
            Top             =   600
            Width           =   3015
         End
         Begin VB.Label UndoLevelsTxt 
            Caption         =   "Levels: 50"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Misc"
         Height          =   1335
         Left            =   120
         TabIndex        =   61
         Top             =   1200
         Width           =   5415
         Begin VB.CheckBox LockingTextures 
            Caption         =   "Textur lo&ck"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            ToolTipText     =   "Locks the texture cordinats to the brush instead of the world."
            Top             =   240
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox TextureAlignment 
            Caption         =   "Align texture to &face"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            ToolTipText     =   "By default aligne the textures to the face of the brush instead of the world."
            Top             =   600
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox GroupWhileIgnore 
            Caption         =   "All&ow grouping/ungrouping while Ignore Groups is checked."
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   960
            Width           =   4815
         End
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Version"
      Height          =   615
      Left            =   120
      TabIndex        =   65
      Top             =   0
      Width           =   5895
      Begin VB.OptionButton Option3 
         Caption         =   "Worldcraft 2.0 - 3.3"
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Hammer 3.4 - 3.5"
         Height          =   195
         Left            =   2160
         TabIndex        =   68
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Hammer 4.0"
         Height          =   195
         Left            =   4080
         TabIndex        =   67
         Top             =   240
         Width           =   1335
      End
   End
   Begin ComctlLib.TabStrip Tabs 
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9763
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&General"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&2D views"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&3D views"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Textures"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "G&UI"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton launch 
      Caption         =   "&Launch"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      ToolTipText     =   "Lunch hammer with the selected options."
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      ToolTipText     =   "Quit the program."
      Top             =   6360
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLunch 
         Caption         =   "&Launch"
      End
      Begin VB.Menu mnuItem1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInport 
         Caption         =   "&Import"
         Index           =   1
         Begin VB.Menu mnuFromTxt 
            Caption         =   "From &file"
            Index           =   1
         End
         Begin VB.Menu mnuFromHammer 
            Caption         =   "From &Hammer"
            Index           =   1
         End
      End
      Begin VB.Menu mnuToTxt 
         Caption         =   "&Save as"
         Index           =   2
      End
      Begin VB.Menu mnuItem2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHammer 
      Caption         =   "H&ammer"
      Begin VB.Menu mnuResetHammer 
         Caption         =   "&Reset"
      End
      Begin VB.Menu mnuResetHammerGui 
         Caption         =   "Reset &GUI"
      End
      Begin VB.Menu mnuClearRecent 
         Caption         =   "&Clear recent File"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'todo
'save/load bars
'clear wad list on loads

Option Explicit

'set exe variable
Private hammer As String

' Note that the TabStrip numbers tabs starting
' with 1 not 0.

' The index of the selected frame.
Private SelectedTab As Integer
Dim value As String

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private lNotepadhWnd As Long

Function IsExecFile(ByVal IsHammerExe As String) As Boolean
    Dim sExt As String

    On Error Resume Next

    ' first of all, check if the file exists
    If Not (GetAttr(IsHammerExe) And vbDirectory) = 0 Then Exit Function

    ' check the last 4 charatcters
    sExt = LCase$(Right$(IsHammerExe, 4))

    ' exe extensions: .exe, .bat, .com, .pif
    IsExecFile = (sExt = ".exe") Or (sExt = ".com") Or (sExt = ".bat") Or (sExt _
        = ".pif")
End Function

Private Sub AnimateModels_Click()
sString = AnimateModels.value
    x = WritePrivateProfileString("3D view", "AnimateModels", sString, HLIFile)
End Sub

Private Sub AutoSelect_Click()
sString = Str(AutoSelect.value)
x = WritePrivateProfileString("2D view", "AutoSelect", sString, HLIFile)
End Sub

Private Sub ClearColor_Click()
On Error GoTo ErrColor
CommonDialog1.Flags = cdlCCRGBInit + cdlCCFullOpen
CommonDialog1.Color = ClearColor.BackColor
CommonDialog1.Action = 3
ClearColor.BackColor = CommonDialog1.Color
sString = ClearColor.BackColor
    x = WritePrivateProfileString("3D view", "ClearColor", sString, HLIFile)
ErrColor:
End Sub


Private Sub ClearColorLabel_Click()
ClearColor_Click
End Sub

Private Sub AddWad_Click()
    On Error GoTo Err1
'brows for textures
    CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Wad file (*.wad)|*.wad|Quake pak file (*.pak)|*.pak"
    CommonDialog1.DialogTitle = "Select texture file"
    CommonDialog1.InitDir = GetSetting("Hammer Launcher", "Setting", "WadDir", CurDir)
    CommonDialog1.ShowOpen
    CommonDialog1.FileName = LCase(CommonDialog1.FileName)
'add it to the list
Dim i As Integer, waddone As Boolean
For i = 0 To (WadList.ListCount - 1)
    If CommonDialog1.FileName = WadList.List(i) Then
        waddone = True
        MsgBox "The .wad/.pak is already on the list."
        i = WadList.ListCount
    End If
Next i
If Not (waddone = True) Then
    WadList.AddItem CommonDialog1.FileName
    SaveSetting "Hammer Launcher", "Setting", "WadDir", CurDir
End If

'save list to hlifile
sString = WadList.ListCount
    x = WritePrivateProfileString("General", "TextureFileCount", sString, HLIFile)

Dim i2 As Integer
For i2 = 0 To (WadList.ListCount - 1)
sString = "TextureFile" + Trim(Str(i2))
    x = WritePrivateProfileString("General", sString, WadList.List(i2), HLIFile)
Next i2

Err1:
End Sub





Private Sub DefaultGrid_LostFocus()
Dim sRetBuf As String, iLenBuf As Integer

On Error GoTo error
If DefaultGrid.Text < 1 Or DefaultGrid.Text > 4096 Then
    GoTo error
End If
x = WritePrivateProfileString("2D view", "Default Grid", DefaultGrid.Text, HLIFile)

GoTo ending

error:

MsgBox "Only natural numbers from 1-4096 are allowed in the grid size feeld."
'todo load saved value
'x = GetPrivateProfileString("2D view", "Default Grid", "64", sRetBuf$, iLenBuf%, HLIFile)
'DefaultGrid.Text = Left$(sRetBuf$, x)
DefaultGrid.Text = 64

ending:

End Sub

Private Sub DrawVertices_Click()
sString = Str(DrawVertices.value)
x = WritePrivateProfileString("2D view", "Draw Vertices", sString, HLIFile)
End Sub

Private Sub FilterTextures_Click()
sString = FilterTextures.value
    x = WritePrivateProfileString("3D view", "FilterTextures", sString, HLIFile)
End Sub

Private Sub GridDots_Click()
sString = GridDots.value
    x = WritePrivateProfileString("2D view", "GridDots", sString, HLIFile)
End Sub

Private Sub Gridhigh1024_Click()
sString = Gridhigh1024.value
    x = WritePrivateProfileString("2D view", "Gridhigh1024", sString, HLIFile)
End Sub

Private Sub Gridhigh64_Click()
sString = Gridhigh64.value
    x = WritePrivateProfileString("2D view", "Gridhigh64", sString, HLIFile)
End Sub

Private Sub GridHighSpec_Change()
Dim sRetBuf As String, iLenBuf As Integer
On Error GoTo error
If GridHighSpec.Text < 0 Or GridHighSpec.Text > 2048 Then
    GoTo error
End If
x = WritePrivateProfileString("2D view", "GridHighSpec", GridHighSpec.Text, HLIFile)
If GridHighSpec.Text > 0 Then
End If
GoTo ending

error:

MsgBox "Only natural numbers betwean 0 & 2048 are allowed."
'todo load saved value
'x = GetPrivateProfileString("2D view", "GridHighSpec", "1", sRetBuf$, iLenBuf%, HLIFile)
'GridHighSpec.Text = Left$(sRetBuf$, x)
GridHighSpec.Text = 64

ending:
End Sub

Private Sub GroupWhileIgnore_Click()
sString = Str(GroupWhileIgnore.value)
x = WritePrivateProfileString("General", "GroupWhileIgnore", sString, HLIFile)
End Sub

Private Sub HideSmallGrid_Click()
sString = Str(KeepCloneGroup.value)
x = WritePrivateProfileString("2D view", "KeepCloneGroup", sString, HLIFile)
End Sub

Private Sub KeepCloneGroup_Click()
sString = Str(KeepCloneGroup.value)
x = WritePrivateProfileString("2D view", "KeepCloneGroup", sString, HLIFile)
End Sub

Private Sub launch_Click()
Dim version As String
Dim sKey As String '***Key in registry
Dim sType As Long '***Value type -- string or number (REG_SZ, REG_DWORD)
Dim regValue As Variant

'what editor is it
If Option1.value = True Then
version = "Software\valve\Valve Hammer Editor\"
'has the hammer path ben set
hammer = GetSetting("Hammer Launcher", "Setting", "hammer", "hammer.exe")
ElseIf Option2.value = True Then
version = "Software\valve\Hammer\"
'has the hammer4 path ben set
hammer = GetSetting("Hammer Launcher", "Setting", "hammer4", "hammer.exe")
Else
version = "Software\valve\Worldcraft\"
'has the wc path ben set
hammer = GetSetting("Hammer Launcher", "Setting", "wc", "hammer.exe")
End If

    '99% of them are dewod's
    sType = REG_DWORD

'2D Views key
    sKey = version + "2D Views"
    CreateNewKey sKey, HKEY_CURRENT_USER
'set all 2d view settings
    SetKeyValue HKEY_CURRENT_USER, sKey, "Default Grid", DefaultGrid.Text, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "RotateConstrain", RotateConstrain.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "Scrollbars", Scrollbars.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "Draw Vertices", DrawVertices.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "WhiteOnBlack", WhiteOnBlack.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "KeepCloneGroup", KeepCloneGroup.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "Usegroupcolors", Usegroupcolors.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "Nudge", Nudge.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "GridHighSpec", GridHighSpec.Text, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "GridDots", GridDots.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "OrientPrimitives", OrientPrimitives.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "AutoSelect", AutoSelect.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "SelectByHandles", SelectByHandles.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "GridIntensity", GridIntensityscroll.value, sType
    If GridHighSpec > 0 Then
        SetKeyValue HKEY_CURRENT_USER, sKey, "GridHigh10", "1", sType
    Else
        SetKeyValue HKEY_CURRENT_USER, sKey, "GridHigh10", "0", sType
    End If
    SetKeyValue HKEY_CURRENT_USER, sKey, "Gridhigh64", Gridhigh64.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "Gridhigh1024", Gridhigh1024.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "HideSmallGrid", HideSmallGrid.value, sType
    
'view's
    sKey = version + "Splitter"
    CreateNewKey sKey, HKEY_CURRENT_USER
'set values
    SetKeyValue HKEY_CURRENT_USER, sKey, "DrawType0,0", view1.ListIndex, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "DrawType0,1", view2.ListIndex, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "DrawType1,0", view3.ListIndex, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "DrawType1,1", view4.ListIndex, sType
    
'3D Views key
    sKey = version + "3D Views"
    CreateNewKey sKey, HKEY_CURRENT_USER
'set all 3d view settings
    SetKeyValue HKEY_CURRENT_USER, sKey, "FilterTextures", FilterTextures.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "AnimateModels", AnimateModels.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "BackPlane", BackPlaneScroll.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "ModelDistance", ModelDistanceScroll.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "UseMouseLook", UseMouseLook.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "Reverse Y", ReverseY.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "ForwardSpeedMax", ForwardSpeedMaxScroll.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "TimeToMaxSpeed", TimeToMaxSpeedScroll.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "ClearColor", ClearColor.BackColor, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "ReverseSelection", ReverseSelection.value, sType
   
    sKey = version + "Barstate-Bar0"
    CreateNewKey sKey, HKEY_CURRENT_USER
    
    
    sKey = version + "Barstate-Bar1"
    CreateNewKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar10"
    CreateNewKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar11"
    CreateNewKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar2"
    CreateNewKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar3"
    CreateNewKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar4"
    CreateNewKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar5"
    CreateNewKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar6"
    CreateNewKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar7"
    CreateNewKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar8"
    CreateNewKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar9"
    CreateNewKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Summary"
    CreateNewKey sKey, HKEY_CURRENT_USER
    
'Configured key
    sKey = version + "Configured"
    CreateNewKey sKey, HKEY_CURRENT_USER
'set this or all other setting will be ignored
    SetKeyValue HKEY_CURRENT_USER, sKey, "Configured", 2, sType

    
'ToDo figure out how this part works
'    sKey = version + "Custom2DColors"
'    CreateNewKey sKey, HKEY_CURRENT_USER

'General key
    sKey = version + "General"
    CreateNewKey sKey, HKEY_CURRENT_USER
'set all General settings
    SetKeyValue HKEY_CURRENT_USER, sKey, "Undo Levels", UndoLevelsScroll.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "Undo Memory Warning", UndoMemoryWarning.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "Independent Windows", IndependentWindows.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "Load Default Positions", LoadDefaultPositions.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "Locking Textures", LockingTextures.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "Texture Alignment", TextureAlignment.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "GroupWhileIgnore", GroupWhileIgnore.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "StretchArches", StretchArches.value, sType
    SetKeyValue HKEY_CURRENT_USER, sKey, "TextureFileCount", WadList.ListCount, sType
If WadList.ListCount > 0 Then


'switch to string
sType = REG_SZ
'add textures
Dim i As Integer
For i = 0 To (WadList.ListCount - 1)
    SetKeyValue HKEY_CURRENT_USER, sKey, "TextureFile" + Trim(Str(i)), WadList.List(i), sType
Next i
End If
    
    sKey = version + "Recent File List"
    CreateNewKey sKey, HKEY_CURRENT_USER
    sKey = version + "Settings"
    CreateNewKey sKey, HKEY_CURRENT_USER
'end of setting files

run:
If IsExecFile(hammer) = True Then
    Call Shell(hammer, vbNormalFocus)
Else
    If Option1.value = True Or Option2.value = True Then
        CommonDialog1.Filter = "Hammer (hammer.exe)|hammer.exe"
        CommonDialog1.DialogTitle = "Select Hammer exe file"
    Else
        CommonDialog1.Filter = "WorldCraft (worldcraft.exe;wc.exe)|worldcraft.exe;wc.exe"
        CommonDialog1.DialogTitle = "Worldcraft exe file"
    End If
On Error GoTo Err1
    CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen
    hammer = CommonDialog1.FileName
    If Option1.value = True Then
        SaveSetting "Hammer Launcher", "Setting", "hammer", hammer
    ElseIf Option2.value = True Then
        SaveSetting "Hammer Launcher", "Setting", "hammer4", hammer
    ElseIf Option3.value = True Then
        SaveSetting "Hammer Launcher", "Setting", "wc2", hammer
    End If
    GoTo run
Err1:
End If

End Sub

Private Sub BackPlaneScroll_Change()
sString = BackPlaneScroll.value
    x = WritePrivateProfileString("3D view", "BackPlane", sString, HLIFile)
BackPlaneScroll_Scroll
End Sub

Private Sub BackPlaneScroll_Scroll()
BackPlaneTxt.Caption = Str(BackPlaneScroll.value)
End Sub


Private Sub LoadDefaultPositions_Click()
sString = LoadDefaultPositions.value
    x = WritePrivateProfileString("General", "Load Default Positions", sString, HLIFile)
End Sub

Private Sub LockingTextures_Click()
sString = Str(LockingTextures.value)
x = WritePrivateProfileString("General", "Locking Textures", sString, HLIFile)
End Sub

Private Sub mnuAbout_Click()
About.Show 1
End Sub

Private Sub mnuClearRecent_Click()
'what version of hammer is it
Dim version As String
If Option1.value = True Then
version = "Software\valve\Valve Hammer Editor\"
ElseIf Option2.value = True Then
version = "Software\valve\Hammer\"
Else
version = "Software\valve\Worldcraft\"
End If

Dim sKey As String
    sKey = version + "Recent File List"
    DeleteKey sKey, HKEY_CURRENT_USER
End Sub

Private Sub mnuEnd_Click()
End
End Sub

Private Sub mnuFromHammer_Click(Index As Integer)

'what version of hammer is it
Dim version As String
If Option1.value = True Then
version = "Software\valve\Valve Hammer Editor\"
ElseIf Option2.value = True Then
version = "Software\valve\Hammer\"
Else
version = "Software\valve\Worldcraft\"
End If

'clear the wad list... the ugly way
start:
If WadList.ListCount > 0 Then GoTo Remove
GoTo ending
Remove:
WadList.RemoveItem (0)
GoTo start
ending:

'start importing the settings
Dim sKey As String '***Key under which to create the value

'2D view
    sKey = version + "2D Views"
'settings
    AutoSelect.value = QueryValue(HKEY_CURRENT_USER, sKey, "AutoSelect")
    Dim GridTemp As Integer
    
    GridTemp = QueryValue(HKEY_CURRENT_USER, sKey, "Default Grid")
    If GridTemp < 1 Or GridTemp > 4096 Then
    DefaultGrid.Text = 64
    Else
    DefaultGrid.Text = GridTemp
    End If
    RotateConstrain.value = QueryValue(HKEY_CURRENT_USER, sKey, "RotateConstrain")
    Scrollbars.value = QueryValue(HKEY_CURRENT_USER, sKey, "Scrollbars")
    DrawVertices.value = QueryValue(HKEY_CURRENT_USER, sKey, "Draw Vertices")
    WhiteOnBlack.value = QueryValue(HKEY_CURRENT_USER, sKey, "WhiteOnBlack")
    KeepCloneGroup.value = QueryValue(HKEY_CURRENT_USER, sKey, "KeepCloneGroup")
    Usegroupcolors.value = QueryValue(HKEY_CURRENT_USER, sKey, "Usegroupcolors")
    Nudge.value = QueryValue(HKEY_CURRENT_USER, sKey, "Nudge")
    GridTemp = QueryValue(HKEY_CURRENT_USER, sKey, "GridHighSpec")
    If GridTemp < 0 Or GridTemp > 2048 Then
    GridHighSpec.Text = 0
    Else
    GridHighSpec.Text = GridTemp
    End If
    GridDots.value = QueryValue(HKEY_CURRENT_USER, sKey, "GridDots")
    OrientPrimitives.value = QueryValue(HKEY_CURRENT_USER, sKey, "OrientPrimitives")
    SelectByHandles.value = QueryValue(HKEY_CURRENT_USER, sKey, "SelectByHandles")
    GridIntensityscroll.value = QueryValue(HKEY_CURRENT_USER, sKey, "GridIntensity")
    Gridhigh64.value = QueryValue(HKEY_CURRENT_USER, sKey, "Gridhigh64")
    Gridhigh1024.value = QueryValue(HKEY_CURRENT_USER, sKey, "Gridhigh1024")
    HideSmallGrid.value = QueryValue(HKEY_CURRENT_USER, sKey, "HideSmallGrid")
    
'Splitter
    sKey = version + "Splitter"
'settings
    view1.ListIndex = QueryValue(HKEY_CURRENT_USER, sKey, "DrawType0,0")
    view2.ListIndex = QueryValue(HKEY_CURRENT_USER, sKey, "DrawType0,1")
    view3.ListIndex = QueryValue(HKEY_CURRENT_USER, sKey, "DrawType1,0")
    view4.ListIndex = QueryValue(HKEY_CURRENT_USER, sKey, "DrawType1,1")
    
'3D Views key
    sKey = version + "3D Views"
'set all 3d view settings
    FilterTextures.value = QueryValue(HKEY_CURRENT_USER, sKey, "FilterTextures")
    AnimateModels.value = QueryValue(HKEY_CURRENT_USER, sKey, "AnimateModels")
    BackPlaneScroll.value = QueryValue(HKEY_CURRENT_USER, sKey, "BackPlane")
    ModelDistanceScroll.value = QueryValue(HKEY_CURRENT_USER, sKey, "ModelDistance")
    UseMouseLook.value = QueryValue(HKEY_CURRENT_USER, sKey, "UseMouseLook")
    ReverseY.value = QueryValue(HKEY_CURRENT_USER, sKey, "Reverse Y")
    ForwardSpeedMaxScroll.value = QueryValue(HKEY_CURRENT_USER, sKey, "ForwardSpeedMax")
    TimeToMaxSpeedScroll.value = QueryValue(HKEY_CURRENT_USER, sKey, "TimeToMaxSpeed")
    ClearColor.BackColor = QueryValue(HKEY_CURRENT_USER, sKey, "ClearColor")
    ReverseSelection.value = QueryValue(HKEY_CURRENT_USER, sKey, "ReverseSelection")
    
'General key
    sKey = version + "General"
'set all General settings
     UndoLevelsScroll.value = QueryValue(HKEY_CURRENT_USER, sKey, "Undo Levels")
     UndoMemoryWarning.value = QueryValue(HKEY_CURRENT_USER, sKey, "Undo Memory Warning")
     IndependentWindows.value = QueryValue(HKEY_CURRENT_USER, sKey, "Independent Windows")
     LoadDefaultPositions.value = QueryValue(HKEY_CURRENT_USER, sKey, "Load Default Positions")
     LockingTextures.value = QueryValue(HKEY_CURRENT_USER, sKey, "Locking Textures")
     TextureAlignment.value = QueryValue(HKEY_CURRENT_USER, sKey, "Texture Alignment")
     GroupWhileIgnore.value = QueryValue(HKEY_CURRENT_USER, sKey, "GroupWhileIgnore")
     StretchArches.value = QueryValue(HKEY_CURRENT_USER, sKey, "StretchArches")

'add textures
If (QueryValue(HKEY_CURRENT_USER, sKey, "TextureFileCount") > 0) Then
Dim i As Integer
Dim i2 As Integer
Dim waddone As Boolean
Dim wad As String
    For i = 0 To (QueryValue(HKEY_CURRENT_USER, sKey, "TextureFileCount") - 1)
        wad = LCase(QueryValue(HKEY_CURRENT_USER, sKey, "TextureFile" + Trim(Str(i))))
        If wad = "texture.wad" Then GoTo EmptyWad
        For i2 = 0 To (WadList.ListCount - 1)
            If wad = WadList.List(i2) Then
            waddone = True
            End If
        Next i2
        If Not (waddone = True) Then
        WadList.AddItem wad
        End If
        waddone = False
EmptyWad:
    Next i
End If

'save wad list to hlifile
sString = WadList.ListCount
    x = WritePrivateProfileString("General", "TextureFileCount", sString, HLIFile)

Dim i3 As Integer
For i3 = 0 To (WadList.ListCount - 1)
sString = "TextureFile" + Trim(Str(i3))
    x = WritePrivateProfileString("General", sString, WadList.List(i3), HLIFile)
Next i3
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Form_Load()
'default sving set

MyDIR = CurDir
Dim i As Integer

    ' Move all the frames to the same position
    ' and make them all invisible.
    For i = 1 To ChoiceFrame.UBound
        ChoiceFrame(i).Move _
            ChoiceFrame(0).Left, _
            ChoiceFrame(0).Top, _
            ChoiceFrame(0).Width, _
            ChoiceFrame(0).Height
        ChoiceFrame(i).Visible = False
    Next i
    
    'also hide the first tab
        ChoiceFrame(0).Visible = False
    
    'Witch tab was last selsect
Dim value As String
value = GetSetting("Hammer Launcher", "Setting", "tab", "1")
If value = "1" Or value = "2" Or value = "3" Or value = "4" Or value = "5" Then
SelectedTab = Int(value)
Else
SelectedTab = 1
End If
    ' Select that tab.
    Tabs.SelectedItem = Tabs.Tabs(SelectedTab)
    ChoiceFrame(SelectedTab - 1).Visible = True
    
'Set GUI
LoadDefaultPositions.Enabled = IndependentWindows.value
If (IndependentWindows.value = 1) Then
    view1.Enabled = False
    view2.Enabled = False
    view3.Enabled = False
    view4.Enabled = False
Else
    view1.Enabled = True
    view2.Enabled = True
    view3.Enabled = True
    view4.Enabled = True
End If
view1.ListIndex = 5
view2.ListIndex = 0
view3.ListIndex = 1
view4.ListIndex = 2

'what was last loaded
HLIFile = GetSetting("Hammer Launcher", "Setting", "HLIFile", "0")
If HLIFile = "0" Then
    mnuFromHammer_Click (1)
    HLIFile = MyDIR + "\Default.hli"
    SaveSetting "Hammer Launcher", "Setting", "HLIFile", HLIFile
Else
    ImportFile (HLIFile)
End If

End Sub

Private Sub ForwardSpeedMaxScroll_Change()
sString = ForwardSpeedMaxScroll.value
    x = WritePrivateProfileString("3D view", "ForwardSpeedMax", sString, HLIFile)
ForwardSpeedMaxScroll_Scroll
End Sub

Private Sub ForwardSpeedMaxScroll_Scroll()
ForwardSpeedMaxTxt.Caption = Str(ForwardSpeedMaxScroll.value)
End Sub

Private Sub GridIntensityscroll_Change()
sString = GridIntensityscroll.value
    x = WritePrivateProfileString("2D view", "GridIntensity", sString, HLIFile)
GridIntensityscroll_scroll
End Sub

Private Sub GridIntensityscroll_scroll()
GridIntensityTxt.Caption = "Intensity:" + Str(GridIntensityscroll.value) + "%"
End Sub

Private Sub IndependentWindows_Click()
    LoadDefaultPositions.Enabled = IndependentWindows.value
If (IndependentWindows.value = 1) Then
    view1.Enabled = False
    view2.Enabled = False
    view3.Enabled = False
    view4.Enabled = False
Else
    view1.Enabled = True
    view2.Enabled = True
    view3.Enabled = True
    view4.Enabled = True
End If

sString = IndependentWindows.value
    x = WritePrivateProfileString("General", "Independent Windows", sString, HLIFile)
End Sub

Private Sub mnuFromTxt_Click(Index As Integer)
Dim file As String

'what file
On Error GoTo ErrIn
    CommonDialog1.InitDir = GetSetting("Hammer Launcher", "Setting", "import", CurDir)
    CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Settings file (*.hli)|*.hli"
    CommonDialog1.DialogTitle = "Select settings file"
    CommonDialog1.ShowOpen
    file$ = CommonDialog1.FileName
    SaveSetting "Hammer Launcher", "Setting", "import", CurDir
    
'Vals
Dim sRetBuf As String, iLenBuf As Integer, sValue As String

'Buffers
sRetBuf$ = String$(256, 0)   '256 null characters
iLenBuf% = Len(sRetBuf$)

HLIFile = file$
SaveSetting "Hammer Launcher", "Setting", "HLIFile", HLIFile
ImportFile (file$)
ErrIn:
End Sub

Private Sub ImportFile(file As String)

'Vals
Dim sSection As String, sRetBuf As String, iLenBuf As Integer, sValue As String

'Buffers
sRetBuf$ = String$(256, 0)   '256 null characters
iLenBuf% = Len(sRetBuf$)


'version
sValue = "File"
x = GetPrivateProfileString(sValue, "version", "3", sRetBuf$, iLenBuf%, file$)
If Int(Left$(sRetBuf$, x)) = 3 Then
Option3.value = False
Option1.value = True
Option2.value = False
ElseIf Int(Left$(sRetBuf$, x)) = 4 Then
Option3.value = False
Option1.value = False
Option2.value = True
Else
Option3.value = True
Option1.value = False
Option2.value = False
End If

'clear the wad list... the ugly way
start:
If WadList.ListCount > 0 Then GoTo Remove
GoTo ending
Remove:
WadList.RemoveItem (0)
GoTo start
ending:

'load settings
sValue = "2D view"
x = GetPrivateProfileString(sValue, "AutoSelect", "0", sRetBuf$, iLenBuf%, file$)
AutoSelect.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "Default Grid", "8", sRetBuf$, iLenBuf%, file$)
DefaultGrid.Text = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "RotateConstrain", "0", sRetBuf$, iLenBuf%, file$)
RotateConstrain.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "Scrollbars", "0", sRetBuf$, iLenBuf%, file$)
Scrollbars.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "Draw Vertices", "0", sRetBuf$, iLenBuf%, file$)
DrawVertices.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "WhiteOnBlack", "0", sRetBuf$, iLenBuf%, file$)
WhiteOnBlack.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "KeepCloneGroup", "0", sRetBuf$, iLenBuf%, file$)
KeepCloneGroup.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "StretchArches", "0", sRetBuf$, iLenBuf%, file$)
StretchArches.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "Usegroupcolors", "0", sRetBuf$, iLenBuf%, file$)
Usegroupcolors.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "Nudge", "0", sRetBuf$, iLenBuf%, file$)
Nudge.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "GridHighSpec", "0", sRetBuf$, iLenBuf%, file$)
GridHighSpec.Text = Left$(sRetBuf$, x)
x = GetPrivateProfileString(sValue, "GridDots", "0", sRetBuf$, iLenBuf%, file$)
GridDots.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "OrientPrimitives", "0", sRetBuf$, iLenBuf%, file$)
OrientPrimitives.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "SelectByHandles", "0", sRetBuf$, iLenBuf%, file$)
SelectByHandles.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "GridIntensity", "0", sRetBuf$, iLenBuf%, file$)
GridIntensityscroll.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "Gridhigh64", "0", sRetBuf$, iLenBuf%, file$)
Gridhigh64.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "Gridhigh1024", "0", sRetBuf$, iLenBuf%, file$)
Gridhigh1024.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "HideSmallGrid", "0", sRetBuf$, iLenBuf%, file$)
HideSmallGrid.value = Int(Left$(sRetBuf$, x))


sValue = "Splitter"
x = GetPrivateProfileString(sValue, "view1", "0", sRetBuf$, iLenBuf%, file$)
view1.ListIndex = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "view2", "0", sRetBuf$, iLenBuf%, file$)
view2.ListIndex = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "view3", "0", sRetBuf$, iLenBuf%, file$)
view3.ListIndex = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "view4", "0", sRetBuf$, iLenBuf%, file$)
view4.ListIndex = Int(Left$(sRetBuf$, x))

sValue = "3D view"
x = GetPrivateProfileString(sValue, "FilterTextures", "0", sRetBuf$, iLenBuf%, file$)
FilterTextures.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "AnimateModels", "0", sRetBuf$, iLenBuf%, file$)
AnimateModels.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "BackPlane", "0", sRetBuf$, iLenBuf%, file$)
BackPlaneScroll.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "ModelDistance", "0", sRetBuf$, iLenBuf%, file$)
ModelDistanceScroll.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "UseMouseLook", "0", sRetBuf$, iLenBuf%, file$)
UseMouseLook.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "Reverse Y", "0", sRetBuf$, iLenBuf%, file$)
ReverseY.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "ForwardSpeedMax", "0", sRetBuf$, iLenBuf%, file$)
ForwardSpeedMaxScroll.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "TimeToMaxSpeed", "0", sRetBuf$, iLenBuf%, file$)
TimeToMaxSpeedScroll.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "ClearColor", "0", sRetBuf$, iLenBuf%, file$)
ClearColor.BackColor = Left$(sRetBuf$, x)
x = GetPrivateProfileString(sValue, "ReverseSelection", "0", sRetBuf$, iLenBuf%, file$)
ReverseSelection.value = Int(Left$(sRetBuf$, x))


sValue = "General"
x = GetPrivateProfileString(sValue, "Undo Levels", "0", sRetBuf$, iLenBuf%, file$)
UndoLevelsScroll.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "Undo Memory Warning", "0", sRetBuf$, iLenBuf%, file$)
UndoMemoryWarning.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "Independent Windows", "0", sRetBuf$, iLenBuf%, file$)
IndependentWindows.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "Load Default Positions", "0", sRetBuf$, iLenBuf%, file$)
LoadDefaultPositions.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "Locking Textures", "0", sRetBuf$, iLenBuf%, file$)
LockingTextures.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "Texture Alignment", "0", sRetBuf$, iLenBuf%, file$)
TextureAlignment.value = Int(Left$(sRetBuf$, x))
x = GetPrivateProfileString(sValue, "GroupWhileIgnore", "0", sRetBuf$, iLenBuf%, file$)
GroupWhileIgnore.value = Int(Left$(sRetBuf$, x))


x = GetPrivateProfileString(sValue, "TextureFileCount", "0", sRetBuf$, iLenBuf%, file$)
If Int(Left$(sRetBuf$, x)) > 0 Then
Dim i As Integer
i = 0
Dim i2 As Integer
Dim waddone As Boolean
Dim wad As String
    For i = 0 To (Int(Left$(sRetBuf$, x)) - 1)
        waddone = False
        wad = "TextureFile" + Trim(Str(i))
    x = GetPrivateProfileString(sValue, wad, "0", sRetBuf$, iLenBuf%, file$)
        wad = LCase(Left$(sRetBuf$, x))
        For i2 = 0 To (WadList.ListCount - 1)
            If wad = WadList.List(i2) Then
            waddone = True
            End If
        Next i2
        If waddone = False Then
        WadList.AddItem wad
        End If
    Next i
End If

End Sub

Private Sub mnuLunch_Click()
launch_Click
End Sub

Private Sub mnuResetHammer_Click()
'what version of hammer is it
Dim version As String
If Option1.value = True Then
version = "Software\valve\Valve Hammer Editor\"
ElseIf Option2.value = True Then
version = "Software\valve\Hammer\"
Else
version = "Software\valve\Worldcraft\"
End If
'delete the info
    DeleteKey version, HKEY_CURRENT_USER
End Sub

Private Sub mnuResetHammerGui_Click()

'what version of hammer is it
Dim version As String
If Option1.value = True Then
version = "Software\valve\Valve Hammer Editor\"
ElseIf Option2.value = True Then
version = "Software\valve\Hammer\"
Else
version = "Software\valve\Worldcraft\"
End If

Dim sKey As String
    sKey = version + "Barstate-Bar1"
    DeleteKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar0"
    DeleteKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar2"
    DeleteKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar3"
    DeleteKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar4"
    DeleteKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar5"
    DeleteKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar6"
    DeleteKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar7"
    DeleteKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar8"
    DeleteKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar9"
    DeleteKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar10"
    DeleteKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Bar11"
    DeleteKey sKey, HKEY_CURRENT_USER
    sKey = version + "Barstate-Summary"
    DeleteKey sKey, HKEY_CURRENT_USER
End Sub

Private Sub mnuToTxt_Click(Index As Integer)
Dim SaveTo As String

On Error GoTo Err1
'where to save
 CommonDialog1.InitDir = GetSetting("Hammer Launcher", "Setting", "export", CurDir)
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Settings file (*.hli)|*.hli"
    CommonDialog1.DialogTitle = "Export to"
    CommonDialog1.Action = 2
    SaveTo$ = CommonDialog1.FileName
SaveSetting "Hammer Launcher", "Setting", "export", CurDir
HLIFile = SaveTo$
SaveSetting "Hammer Launcher", "Setting", "HLIFile", HLIFile

If Option1.value = True Then
    x = WritePrivateProfileString("File", "version", "3", SaveTo)
ElseIf Option2.value = True Then
    x = WritePrivateProfileString("File", "version", "4", SaveTo)
Else
    x = WritePrivateProfileString("File", "version", "2", SaveTo)
End If

'2D view
sSection$ = "2D view"
sString$ = AutoSelect.value
    x = WritePrivateProfileString(sSection$, "AutoSelect", sString$, SaveTo$)
    x = WritePrivateProfileString(sSection$, "Default Grid", DefaultGrid.Text, SaveTo$)
sString$ = RotateConstrain.value
    x = WritePrivateProfileString(sSection$, "RotateConstrain", sString$, SaveTo$)
sString$ = Scrollbars.value
    x = WritePrivateProfileString(sSection$, "Scrollbars", sString$, SaveTo$)
sString$ = DrawVertices.value
    x = WritePrivateProfileString(sSection$, "Draw Vertices", sString$, SaveTo$)
sString$ = WhiteOnBlack.value
    x = WritePrivateProfileString(sSection$, "WhiteOnBlack", sString$, SaveTo$)
sString$ = KeepCloneGroup.value
    x = WritePrivateProfileString(sSection$, "KeepCloneGroup", sString$, SaveTo$)
sString$ = StretchArches.value
    x = WritePrivateProfileString(sSection$, "StretchArches", sString$, SaveTo$)
sString$ = Usegroupcolors.value
    x = WritePrivateProfileString(sSection$, "Usegroupcolors", sString$, SaveTo$)
sString$ = Nudge.value
    x = WritePrivateProfileString(sSection$, "Nudge", sString$, SaveTo$)
    x = WritePrivateProfileString(sSection$, "GridHighSpec", GridHighSpec.Text, SaveTo$)
sString$ = GridDots.value
    x = WritePrivateProfileString(sSection$, "GridDots", sString$, SaveTo$)
sString$ = OrientPrimitives.value
    x = WritePrivateProfileString(sSection$, "OrientPrimitives", sString$, SaveTo$)
sString$ = SelectByHandles.value
    x = WritePrivateProfileString(sSection$, "SelectByHandles", sString$, SaveTo$)
sString$ = GridIntensityscroll.value
    x = WritePrivateProfileString(sSection$, "GridIntensity", sString$, SaveTo$)
sString$ = Gridhigh64.value
    x = WritePrivateProfileString(sSection$, "Gridhigh64", sString$, SaveTo$)
sString$ = Gridhigh1024.value
    x = WritePrivateProfileString(sSection$, "Gridhigh1024", sString$, SaveTo$)
sString$ = HideSmallGrid.value
    x = WritePrivateProfileString(sSection$, "HideSmallGrid", sString$, SaveTo$)
    
    '3D Views
sSection$ = "Splitter"
sString$ = view1.ListIndex
    x = WritePrivateProfileString(sSection$, "view1", sString$, SaveTo$)
sString$ = view2.ListIndex
    x = WritePrivateProfileString(sSection$, "view2", sString$, SaveTo$)
sString$ = view3.ListIndex
    x = WritePrivateProfileString(sSection$, "view3", sString$, SaveTo$)
sString$ = view4.ListIndex
    x = WritePrivateProfileString(sSection$, "view4", sString$, SaveTo$)
        
    
    
    '3D Views key
sSection$ = "3D view"
sString$ = FilterTextures.value
    x = WritePrivateProfileString(sSection$, "FilterTextures", sString$, SaveTo$)
sString$ = AnimateModels.value
    x = WritePrivateProfileString(sSection$, "AnimateModels", sString$, SaveTo$)
sString$ = BackPlaneScroll.value
    x = WritePrivateProfileString(sSection$, "BackPlane", sString$, SaveTo$)
sString$ = ModelDistanceScroll.value
    x = WritePrivateProfileString(sSection$, "ModelDistance", sString$, SaveTo$)
sString$ = UseMouseLook.value
    x = WritePrivateProfileString(sSection$, "UseMouseLook", sString$, SaveTo$)
sString$ = ReverseY.value
    x = WritePrivateProfileString(sSection$, "Reverse Y", sString$, SaveTo$)
sString$ = ForwardSpeedMaxScroll.value
    x = WritePrivateProfileString(sSection$, "ForwardSpeedMax", sString$, SaveTo$)
sString$ = TimeToMaxSpeedScroll.value
    x = WritePrivateProfileString(sSection$, "TimeToMaxSpeed", sString$, SaveTo$)
sString$ = ClearColor.BackColor
    x = WritePrivateProfileString(sSection$, "ClearColor", sString$, SaveTo$)
sString$ = ReverseSelection.value
    x = WritePrivateProfileString(sSection$, "ReverseSelection", sString$, SaveTo$)
    
    'General key
sSection$ = "General"
sString$ = UndoLevelsScroll.value
    x = WritePrivateProfileString(sSection$, "Undo Levels", sString$, SaveTo$)
sString$ = UndoMemoryWarning.value
    x = WritePrivateProfileString(sSection$, "Undo Memory Warning", sString$, SaveTo$)
sString$ = IndependentWindows.value
    x = WritePrivateProfileString(sSection$, "Independent Windows", sString$, SaveTo$)
sString$ = LoadDefaultPositions.value
    x = WritePrivateProfileString(sSection$, "Load Default Positions", sString$, SaveTo$)
sString$ = LockingTextures.value
    x = WritePrivateProfileString(sSection$, "Locking Textures", sString$, SaveTo$)
sString$ = TextureAlignment.value
    x = WritePrivateProfileString(sSection$, "Texture Alignment", sString$, SaveTo$)
sString$ = GroupWhileIgnore.value
    x = WritePrivateProfileString(sSection$, "GroupWhileIgnore", sString$, SaveTo$)
sString$ = WadList.ListCount
    x = WritePrivateProfileString(sSection$, "TextureFileCount", sString$, SaveTo$)
    
Dim i As Integer
For i = 0 To (WadList.ListCount - 1)
sString$ = "TextureFile" + Trim(Str(i))
    x = WritePrivateProfileString(sSection$, sString$, WadList.List(i), SaveTo$)
Next i
Err1:
End Sub

Private Sub ModelDistanceScroll_Change()
sString = ModelDistanceScroll.value
    x = WritePrivateProfileString("3D view", "ModelDistance", sString, HLIFile)
ModelDistanceScroll_Scroll
End Sub

Private Sub ModelDistanceScroll_Scroll()
ModelDistanceTxt.Caption = Str(ModelDistanceScroll.value)
End Sub

Private Sub Nudge_Click()
sString = Str(Nudge.value)
x = WritePrivateProfileString("2D view", "Nudge", sString, HLIFile)
End Sub

Private Sub Option1_Click()
AnimateModels.Enabled = True
Label8.Enabled = True
ModelDistanceTxt.Enabled = True
ModelDistanceScroll.Enabled = True
ClearColorLabel.Enabled = True
ClearColor.Enabled = True
x = WritePrivateProfileString("File", "version", "3", HLIFile)
End Sub

Private Sub Option2_Click()
AnimateModels.Enabled = True
Label8.Enabled = True
ModelDistanceTxt.Enabled = True
ModelDistanceScroll.Enabled = True
ClearColorLabel.Enabled = False
ClearColor.Enabled = False
x = WritePrivateProfileString("File", "version", "4", HLIFile)
End Sub

Private Sub Option3_Click()
AnimateModels.Enabled = False
Label8.Enabled = False
ModelDistanceTxt.Enabled = False
ModelDistanceScroll.Enabled = False
ClearColorLabel.Enabled = False
ClearColor.Enabled = False

x = WritePrivateProfileString("File", "version", "2", HLIFile)
End Sub



Private Sub OrientPrimitives_Click()
sString = Str(OrientPrimitives.value)
x = WritePrivateProfileString("2D view", "OrientPrimitives", sString, HLIFile)
End Sub

Private Sub ReverseSelection_Click()
sString = ReverseSelection.value
    x = WritePrivateProfileString("3D view", "ReverseSelection", sString, HLIFile)
End Sub

Private Sub ReverseY_Click()
sString = ReverseY.value
    x = WritePrivateProfileString(sSection$, "Reverse Y", sString, HLIFile)
End Sub

Private Sub RotateConstrain_Click()
sString = Str(RotateConstrain.value)
x = WritePrivateProfileString("2D view", "RotateConstrain", sString, HLIFile)
End Sub

Private Sub Scrollbars_Click()
sString = Str(Scrollbars.value)
x = WritePrivateProfileString("2D view", "Scrollbars", sString, HLIFile)
End Sub

Private Sub SelectByHandles_Click()
sString = Str(SelectByHandles.value)
x = WritePrivateProfileString("General", "SelectByHandles", sString, HLIFile)
End Sub

Private Sub StretchArches_Click()
sString = Str(StretchArches.value)
x = WritePrivateProfileString("2D view", "StretchArches", sString, HLIFile)
End Sub

Private Sub TabS_Click()
    ChoiceFrame(SelectedTab - 1).Visible = False
    SelectedTab = Tabs.SelectedItem.Index
    ChoiceFrame(SelectedTab - 1).Visible = True
    
'save selecktion to the reg
    SaveSetting "Hammer Launcher", "Setting", "tab", SelectedTab
End Sub

Private Sub TextureAlignment_Click()
sString = Str(TextureAlignment.value)
x = WritePrivateProfileString("General", "Texture Alignment", sString, HLIFile)
End Sub

Private Sub TimeToMaxSpeedScroll_Change()
sString = Str(TimeToMaxSpeedScroll.value)
x = WritePrivateProfileString("3D view", "TimeToMaxSpeed", sString, HLIFile)
TimeToMaxSpeedScroll_scroll
End Sub

Private Sub TimeToMaxSpeedScroll_scroll()
TimeToMaxSpeedTxt.Caption = Str(TimeToMaxSpeedScroll.value / 1000) + " sec."
End Sub

Private Sub UndoLevelsScroll_Change()
sString = Str(UndoLevelsScroll.value)
x = WritePrivateProfileString("General", "Undo Levels", sString, HLIFile)
UndoLevelsScroll_Scroll
End Sub

Private Sub UndoLevelsScroll_Scroll()
UndoLevelsTxt.Caption = "levels:" + Str(UndoLevelsScroll.value)
End Sub

Private Sub UndoMemoryWarning_Click()
sString = Str(UndoMemoryWarning.value)
x = WritePrivateProfileString("General", "Undo Memory Warning", sString, HLIFile)
End Sub

Private Sub Usegroupcolors_Click()
sString = Str(Usegroupcolors.value)
x = WritePrivateProfileString("2D view", "Usegroupcolors", sString, HLIFile)
End Sub

Private Sub UseMouseLook_Click()
sString = UseMouseLook.value
    x = WritePrivateProfileString("3D view", "UseMouseLook", sString, HLIFile)
End Sub

Private Sub view1_Click()
If view1.ListIndex = 3 Then
Picture1.Picture = ImgWire.Picture
ElseIf view1.ListIndex = 4 Then
Picture1.Picture = ImgSolid.Picture
ElseIf view1.ListIndex = 5 Then
Picture1.Picture = ImgTex.Picture
Else
Picture1.Picture = Img2D.Picture
End If

sString = view1.ListIndex
    x = WritePrivateProfileString("Splitter", "view1", sString, HLIFile)
End Sub

Private Sub view2_Click()
If view2.ListIndex = 3 Then
Picture2.Picture = ImgWire.Picture
ElseIf view2.ListIndex = 4 Then
Picture2.Picture = ImgSolid.Picture
ElseIf view2.ListIndex = 5 Then
Picture2.Picture = ImgTex.Picture
Else
Picture2.Picture = Img2D.Picture
End If

sString = view2.ListIndex
    x = WritePrivateProfileString("Splitter", "view2", sString, HLIFile)
End Sub

Private Sub view3_Click()
If view3.ListIndex = 3 Then
Picture3.Picture = ImgWire.Picture
ElseIf view3.ListIndex = 4 Then
Picture3.Picture = ImgSolid.Picture
ElseIf view3.ListIndex = 5 Then
Picture3.Picture = ImgTex.Picture
Else
Picture3.Picture = Img2D.Picture
End If

sString = view3.ListIndex
    x = WritePrivateProfileString("Splitter", "view3", sString, HLIFile)
End Sub

Private Sub view4_Click()
If view4.ListIndex = 3 Then
Picture4.Picture = ImgWire.Picture
ElseIf view4.ListIndex = 4 Then
Picture4.Picture = ImgSolid.Picture
ElseIf view4.ListIndex = 5 Then
Picture4.Picture = ImgTex.Picture
Else
Picture4.Picture = Img2D.Picture
End If

sString = view4.ListIndex
    x = WritePrivateProfileString("Splitter", "view4", sString, HLIFile)
End Sub

Private Sub WadList_Click()
    If WadList.ListIndex > -1 Then
    WadRemove.Enabled = True
    End If
End Sub

Private Sub WadRemove_Click()
WadList.RemoveItem (WadList.ListIndex)
If WadList.ListIndex > -1 Then
    WadRemove.Enabled = True
Else
    WadRemove.Enabled = False
End If

'save list to hlifile
sString = WadList.ListCount
    x = WritePrivateProfileString("General", "TextureFileCount", sString, HLIFile)

Dim i As Integer
For i = 0 To (WadList.ListCount - 1)
sString = "TextureFile" + Trim(Str(i))
    x = WritePrivateProfileString("General", sString, WadList.List(i), HLIFile)
Next i

End Sub

Private Sub todo()
MsgBox "Still to come."
End Sub

Private Sub WhiteOnBlack_Click()
sString = Str(WhiteOnBlack.value)
x = WritePrivateProfileString("2D view", "WhiteOnBlack", sString, HLIFile)
End Sub
