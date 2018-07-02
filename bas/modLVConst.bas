Attribute VB_Name = "modLVConst"
Option Explicit

' (Win95)
Public Const WC_LISTVIEWA = "SysListView32"
Public Const WC_LISTVIEW = WC_LISTVIEWA

Const LVM_FIRST = &H1000

'ListView Styles(LVS_)
Public Const LVS_ICON = &H0
Public Const LVS_REPORT = &H1
Public Const LVS_SMALLICON = &H2
Public Const LVS_LIST = &H3
Public Const LVS_TYPEMASK = &H3
Public Const LVS_SINGLESEL = &H4
Public Const LVS_SHOWSELALWAYS = &H8
Public Const LVS_SORTASCENDING = &H10
Public Const LVS_SORTDESCENDING = &H20
Public Const LVS_SHAREIMAGELISTS = &H40
Public Const LVS_NOLABELWRAP = &H80
Public Const LVS_AUTOARRANGE = &H100
Public Const LVS_EDITLABELS = &H200
Public Const LVS_OWNERDATA = &H1000         'IE 3+ only
Public Const LVS_NOSCROLL = &H2000

Public Const LVS_TYPESTYLEMASK = &HFC00

Public Const LVS_ALIGNTOP = &H0
Public Const LVS_ALIGNLEFT = &H800
Public Const LVS_ALIGNMASK = &HC00

Public Const LVS_OWNERDRAWFIXED = &H400
Public Const LVS_NOCOLUMNHEADER = &H4000
Public Const LVS_NOSORTHEADER = &H8000

'------------------------------------------------------------------------
'ListView Messages(LVM_)(Generic)

Public Const LVM_GETBKCOLOR = (LVM_FIRST + 0)
Public Const LVM_SETBKCOLOR = (LVM_FIRST + 1)
Public Const LVM_GETIMAGELIST = (LVM_FIRST + 2)
Public Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
Public Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)

Public Const LVM_DELETEITEM = (LVM_FIRST + 8)
Public Const LVM_DELETEALLITEMS = (LVM_FIRST + 9)
Public Const LVM_GETCALLBACKMASK = (LVM_FIRST + 10)
Public Const LVM_SETCALLBACKMASK = (LVM_FIRST + 11)
Public Const LVM_GETNEXTITEM = (LVM_FIRST + 12)

Public Const LVM_SETITEMPOSITION = (LVM_FIRST + 15)
Public Const LVM_GETITEMPOSITION = (LVM_FIRST + 16)

Public Const LVM_HITTEST = (LVM_FIRST + 18)
Public Const LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
Public Const LVM_SCROLL = (LVM_FIRST + 20)
Public Const LVM_REDRAWITEMS = (LVM_FIRST + 21)
Public Const LVM_ARRANGE = (LVM_FIRST + 22)

Public Const LVM_GETEDITCONTROL = (LVM_FIRST + 24)

Public Const LVM_DELETECOLUMN = (LVM_FIRST + 28)
Public Const LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)

Public Const LVM_GETHEADER = (LVM_FIRST + 31)   'IE 3+ only

Public Const LVM_CREATEDRAGIMAGE = (LVM_FIRST + 33)
Public Const LVM_GETVIEWRECT = (LVM_FIRST + 34)
Public Const LVM_GETTEXTCOLOR = (LVM_FIRST + 35)
Public Const LVM_SETTEXTCOLOR = (LVM_FIRST + 36)
Public Const LVM_GETTEXTBKCOLOR = (LVM_FIRST + 37)
Public Const LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)
Public Const LVM_GETTOPINDEX = (LVM_FIRST + 39)
Public Const LVM_GETCOUNTPERPAGE = (LVM_FIRST + 40)
Public Const LVM_GETORIGIN = (LVM_FIRST + 41)
Public Const LVM_UPDATE = (LVM_FIRST + 42)
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE = (LVM_FIRST + 44)
Public Const LVM_SETITEMCOUNT = (LVM_FIRST + 47)
Public Const LVM_SORTITEMS = (LVM_FIRST + 48)
Public Const LVM_SETITEMPOSITION32 = (LVM_FIRST + 49)
Public Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
Public Const LVM_GETITEMSPACING = (LVM_FIRST + 51)

Public Const LVM_SETICONSPACING = (LVM_FIRST + 53)  'IE 3+ only

Public Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Public Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
Public Const LVM_SETCOLUMNORDERARRAY = (LVM_FIRST + 58)
Public Const LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
Public Const LVM_SETHOTITEM = (LVM_FIRST + 60)
Public Const LVM_GETHOTITEM = (LVM_FIRST + 61)
Public Const LVM_SETHOTCURSOR = (LVM_FIRST + 62)
Public Const LVM_GETHOTCURSOR = (LVM_FIRST + 63)
Public Const LVM_APPROXIMATEVIEWRECT = (LVM_FIRST + 64)
Public Const LVM_SETWORKAREA = (LVM_FIRST + 65)

Public Const LVM_GETSELECTIONMARK = (LVM_FIRST + 66)  'Win32 and IE 4 only
Public Const LVM_SETSELECTIONMARK = (LVM_FIRST + 67)  'Win32 and IE 4 only
Public Const LVM_GETWORKAREA = (LVM_FIRST + 70)       'Win32 and IE 4 only
Public Const LVM_SETHOVERTIME = (LVM_FIRST + 71)      'Win32 and IE 4 only
Public Const LVM_GETHOVERTIME = (LVM_FIRST + 72)      'Win32 and IE 4 only 

'------------------------------------------------------------------------
'ListView Messages(LVM_)(Win95 - specific)

Public Const LVM_GETITEM = (LVM_FIRST + 5)
Public Const LVM_SETITEM = (LVM_FIRST + 6)

Public Const LVM_INSERTITEMA = (LVM_FIRST + 7)
Public Const LVM_INSERTITEM = LVM_INSERTITEMA

Public Const LVM_FINDITEMA = (LVM_FIRST + 13)
Public Const LVM_FINDITEM = LVM_FINDITEMA

Public Const LVM_GETSTRINGWIDTHA = (LVM_FIRST + 17)
Public Const LVM_GETSTRINGWIDTH = LVM_GETSTRINGWIDTHA

Public Const LVM_EDITLABELA = (LVM_FIRST + 23)
Public Const LVM_EDITLABEL = LVM_EDITLABELA

Public Const LVM_GETCOLUMNA = (LVM_FIRST + 25)
Public Const LVM_GETCOLUMN = LVM_GETCOLUMNA

Public Const LVM_SETCOLUMNA = (LVM_FIRST + 26)
Public Const LVM_SETCOLUMN = LVM_SETCOLUMNA

Public Const LVM_INSERTCOLUMNA = (LVM_FIRST + 27)
Public Const LVM_INSERTCOLUMN = LVM_INSERTCOLUMNA

Public Const LVM_GETITEMTEXTA = (LVM_FIRST + 45)
Public Const LVM_GETITEMTEXT = LVM_GETITEMTEXTA

Public Const LVM_SETITEMTEXTA = (LVM_FIRST + 46)
Public Const LVM_SETITEMTEXT = LVM_SETITEMTEXTA

Public Const LVM_GETISEARCHSTRINGA = (LVM_FIRST + 52)
Public Const LVM_GETISEARCHSTRING = LVM_GETISEARCHSTRINGA

Public Const LVM_SETBKIMAGEA = (LVM_FIRST + 68)         'Win32 and IE 4 only
Public Const LVM_GETBKIMAGEA = (LVM_FIRST + 69)         'Win32 and IE 4 only
'Public Const LVBKIMAGE = LVBKIMAGEA                     'Win32 and IE 4 only
'Public Const LPLVBKIMAGE = LPLVBKIMAGEA                 'Win32 and IE 4 only
Public Const LVM_SETBKIMAGE = LVM_SETBKIMAGEA           'Win32 and IE 4 only
Public Const LVM_GETBKIMAGE = LVM_GETBKIMAGEA           'Win32 and IE 4 only

'------------------------------------------------------------------------
'ListView Messages(LVM_)(Unicode - specific)

'Public Const LVM_GETITEM = (LVM_FIRST + 75)
'Public Const LVM_SETITEM = (LVM_FIRST + 76)
'
'Public Const LVM_INSERTITEMW = (LVM_FIRST + 77)
'Public Const LVM_INSERTITEM = LVM_INSERTITEMW
'
'Public Const LVM_FINDITEMW = (LVM_FIRST + 83)
'Public Const LVM_FINDITEM = LVM_FINDITEMW
'
'Public Const LVM_GETSTRINGWIDTHW = (LVM_FIRST + 87)
'Public Const LVM_GETSTRINGWIDTH = LVM_GETSTRINGWIDTHW
'
'Public Const LVM_EDITLABELW = (LVM_FIRST + 118)
'Public Const LVM_EDITLABEL = LVM_EDITLABELW
'
'Public Const LVM_GETCOLUMNW = (LVM_FIRST + 95)
'Public Const LVM_GETCOLUMN = LVM_GETCOLUMNW
'
'Public Const LVM_SETCOLUMNW = (LVM_FIRST + 96)
'Public Const LVM_SETCOLUMN = LVM_SETCOLUMNW
'
'Public Const LVM_INSERTCOLUMNW = (LVM_FIRST + 97)
'Public Const LVM_INSERTCOLUMN = LVM_INSERTCOLUMNW
'
'Public Const LVM_GETITEMTEXTW = (LVM_FIRST + 115)
'Public Const LVM_GETITEMTEXT = LVM_GETITEMTEXTW
'
'Public Const LVM_SETITEMTEXTW = (LVM_FIRST + 116)
'Public Const LVM_SETITEMTEXT = LVM_SETITEMTEXTW
'
'Public Const LVM_GETISEARCHSTRINGW = (LVM_FIRST + 117)
'Public Const LVM_GETISEARCHSTRING = LVM_GETISEARCHSTRINGW
'
'Public Const LVM_GETBKIMAGEW = (LVM_FIRST + 139)        'Win32 and IE 4 only
'Public Const LVM_SETBKIMAGEW = (LVM_FIRST + 138)        'Win32 and IE 4 only
'Public Const LVBKIMAGE = LVBKIMAGEW                     'Win32 and IE 4 only
'Public Const LPLVBKIMAGE = LPLVBKIMAGEW                 'Win32 and IE 4 only
'Public Const LVM_SETBKIMAGE = LVM_SETBKIMAGEW           'Win32 and IE 4 only
'Public Const LVM_GETBKIMAGE = LVM_GETBKIMAGEW           'Win32 and IE 4 only

'------------------------------------------------------------------------
'ListView Extended Style Messages (LVS_EX_) (Win95-specific)

Public Const LVS_EX_GRIDLINES = &H1
Public Const LVS_EX_SUBITEMIMAGES = &H2
Public Const LVS_EX_CHECKBOXES = &H4
Public Const LVS_EX_TRACKSELECT = &H8
Public Const LVS_EX_HEADERDRAGDROP = &H10
Public Const LVS_EX_FULLROWSELECT = &H20       'applies to report mode only
Public Const LVS_EX_ONECLICKACTIVATE = &H40
Public Const LVS_EX_TWOCLICKACTIVATE = &H80
Public Const LVS_EX_FLATSB = &H100             'cannot be cleared - Win32 & IE4 only
Public Const LVS_EX_REGIONAL = &H200           'Win32 & IE4 only
Public Const LVS_EX_INFOTIP = &H400            'listview does InfoTips for you - Win32 & IE4 only

'------------------------------------------------------------------------
'ListView Set Image List Messages (LVSIL_)

Public Const LVSIL_NORMAL = 0
Public Const LVSIL_SMALL = 1
Public Const LVSIL_STATE = 2

'------------------------------------------------------------------------
'ListView Item Format Messages (LVIF_)

Public Const LVIF_TEXT = &H1
Public Const LVIF_IMAGE = &H2
Public Const LVIF_PARAM = &H4
Public Const LVIF_STATE = &H8
Public Const LVIF_INDENT = &H10           'IE 3+ only
Public Const LVIF_NORECOMPUTE = &H800     'IE 3+ only
Public Const LVIF_DI_SETITEM = &H1000

'------------------------------------------------------------------------
'ListView Item State Messages (LVIS_)

Public Const LVIS_FOCUSED = &H1
Public Const LVIS_SELECTED = &H2
Public Const LVIS_CUT = &H4
Public Const LVIS_DROPHILITED = &H8

Public Const LVIS_OVERLAYMASK = &HF00
Public Const LVIS_STATEIMAGEMASK = &HF000

'------------------------------------------------------------------------
'ListView Item Definitions (LVITEM) (Win95)

'Public Const LVITEM = LVITEMA
'Public Const LPLVITEM = LPLVITEMA
'Public Const LV_ITEMA = LVITEMA       'IE 3+ only
'Public Const tagLVITEMA = LV_ITEMA


'ListView Item Definitions (LVITEM) (Unicode)

'Public Const LVITEM = LVITEMW
'Public Const LPLVITEM = LPLVITEMW  'Unicode (NT)
'Public Const LV_ITEM = LVITEM      'IE 3+ only
'Public Const tagLVITEMW = LV_ITEMW

'------------------------------------------------------------------------
'ListView -Misc.Messages

'Public Const INDEXTOSTATEIMAGEMASK(i) ((i) << 12)
Public Const I_INDENTCALLBACK = (-1)               'IE 3+ only
'Public Const LPSTR_TEXTCALLBACKW = ((LPWSTR) - 1&) 'Unicode (NT)
'Public Const LPSTR_TEXTCALLBACKA = ((LPSTR) - 1&)  'win95

'Public Const LPSTR_TEXTCALLBACK = LPSTR_TEXTCALLBACKW 'Unicode (NT)
'Public Const LPSTR_TEXTCALLBACK = LPSTR_TEXTCALLBACKA 'win95

'------------------------------------------------------------------------
'ListView Notification Item Messages (LVNI_)

Public Const LVNI_ALL = &H0
Public Const LVNI_FOCUSED = &H1
Public Const LVNI_SELECTED = &H2
Public Const LVNI_CUT = &H4
Public Const LVNI_DROPHILITED = &H8

Public Const LVNI_ABOVE = &H100
Public Const LVNI_BELOW = &H200
Public Const LVNI_TOLEFT = &H400
Public Const LVNI_TORIGHT = &H800

'------------------------------------------------------------------------
'ListView Find Item Messages (LVFI_) (Generic)

Public Const LVFI_PARAM = &H1
Public Const LVFI_STRING = &H2
Public Const LVFI_PARTIAL = &H8
Public Const LVFI_WRAP = &H20
Public Const LVFI_NEARESTXY = &H40

'Public Const LV_FINDINFO = LVFINDINFO

'------------------------------------------------------------------------
'ListView Find Item Messages (LVFI_) (Win95)

'Public Const LV_FINDINFOA = LVFINDINFOA
'Public Const LV_FINDINFOA = LVFINDINFOA     'IE 3+ only
'Public Const tagLVFINDINFOA = LV_FINDINFOA
'Public Const LVFINDINFOA = LV_FINDINFOA
'Public Const LVFINDINFO = LVFINDINFOA

'------------------------------------------------------------------------
'ListView Find Item Messages (LVFI_) (Unicode)

'Public Const LV_FINDINFOW = LVFINDINFOW
'Public Const LV_FINDINFOW = LVFINDINFOW     'IE 3+ only
'Public Const tagLVFINDINFOW = LV_FINDINFOW
'Public Const LVFINDINFOW = LV_FINDINFOW
'Public Const LVFINDINFO = LVFINDINFOW

'------------------------------------------------------------------------
'ListView Find ItemRect Messages (LVIR_)

Public Const LVIR_BOUNDS = 0
Public Const LVIR_ICON = 1
Public Const LVIR_LABEL = 2
Public Const LVIR_SELECTBOUNDS = 3
'ListView Hit Test Messages (LVHT_)
Public Const LVHT_NOWHERE = &H1
Public Const LVHT_ONITEMICON = &H2
Public Const LVHT_ONITEMLABEL = &H4
Public Const LVHT_ONITEMSTATEICON = &H8
Public Const LVHT_ONITEM = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)

Public Const LVHT_ABOVE = &H8
Public Const LVHT_BELOW = &H10
Public Const LVHT_TORIGHT = &H20
Public Const LVHT_TOLEFT = &H40

'Public Const LV_HITTESTINFO = LVHITTESTINFO       'IE 3+ only
'Public Const tagLVHITTESTINFO = LV_HITTESTINFO
'Public Const LVHITTESTINFO = LV_HITTESTINFO

'------------------------------------------------------------------------
'ListView Arrange Messages (LVA_)

Public Const LVA_DEFAULT = &H0
Public Const LVA_ALIGNLEFT = &H1
Public Const LVA_ALIGNTOP = &H2
Public Const LVA_SNAPTOGRID = &H5

'------------------------------------------------------------------------
'ListView Column Messages (LVC_) (Generic)

'Public Const LV_COLUMN = LVCOLUMN       'IE 3+ only

'------------------------------------------------------------------------
'ListView Column Messages (LVC_) (Win95)

'Public Const LV_COLUMNA = LVCOLUMNA               'IE 3+ only
'Public Const tagLVCOLUMNA = LV_COLUMNA
'Public Const LVCOLUMNA = LV_COLUMNA
'Public Const LVCOLUMN = LVCOLUMNA
'Public Const LPLVCOLUMN = LPLVCOLUMNA

'------------------------------------------------------------------------
'ListView Column Messages (LVC_) (Unicode)

'Public Const LV_COLUMNW = LVCOLUMNW       'IE 3+ only
'Public Const tagLVCOLUMNW = LV_COLUMNW
'Public Const LVCOLUMNW = LV_COLUMNW
'Public Const LVCOLUMN = LVCOLUMNW
'Public Const LPLVCOLUMN = LPLVCOLUMNW

'------------------------------------------------------------------------
'ListView Column Flag Messages (LVCF_) (LVC.mask)

Public Const LVCF_FMT = &H1
Public Const LVCF_WIDTH = &H2
Public Const LVCF_TEXT = &H4
Public Const LVCF_SUBITEM = &H8
Public Const LVCF_IMAGE = &H10      'IE 3+ only
Public Const LVCF_ORDER = &H20      'IE 3+ only

'------------------------------------------------------------------------
'ListView Column Format Messages (LVCFMT_) (LVC.fmt)

Public Const LVCFMT_LEFT = &H0
Public Const LVCFMT_RIGHT = &H1
Public Const LVCFMT_CENTER = &H2
Public Const LVCFMT_JUSTIFYMASK = &H3
Public Const LVCFMT_IMAGE = &H800               'IE 3+ only
Public Const LVCFMT_BITMAP_ON_RIGHT = &H1000    'IE 3+ only
Public Const LVCFMT_COL_HAS_IMAGES = &H8000     'IE 4 only

'------------------------------------------------------------------------
'ListView Set Column Width Messages (LVSCW_)

Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2

'------------------------------------------------------------------------
'ListView Background Image Flags (LVBKIF_)

Public Const LVBKIF_SOURCE_NONE = &H0       'Win32 and IE 4 only
Public Const LVBKIF_SOURCE_HBITMAP = &H1    'Win32 and IE 4 only
Public Const LVBKIF_SOURCE_URL = &H2        'Win32 and IE 4 only
Public Const LVBKIF_SOURCE_MASK = &H3       'Win32 and IE 4 only
Public Const LVBKIF_STYLE_NORMAL = &H0      'Win32 and IE 4 only
Public Const LVBKIF_STYLE_TILE = &H10       'Win32 and IE 4 only
Public Const LVBKIF_STYLE_MASK = &H10       'Win32 and IE 4 only

'------------------------------------------------------------------------
'ListView Notification Messages (LVN_) (Generic)

'Public Const LVN_ITEMCHANGING = (LVN_FIRST - 0)
'Public Const LVN_ITEMCHANGED = (LVN_FIRST - 1)
'Public Const LVN_INSERTITEM = (LVN_FIRST - 2)
'Public Const LVN_DELETEITEM = (LVN_FIRST - 3)
'Public Const LVN_DELETEALLITEMS = (LVN_FIRST - 4)
'
'Public Const LVN_COLUMNCLICK = (LVN_FIRST - 8)
'Public Const LVN_BEGINDRAG = (LVN_FIRST - 9)
'Public Const LVN_BEGINRDRAG = (LVN_FIRST - 11)
'
'Public Const LVN_ODCACHEHINT = (LVN_FIRST - 13)         'IE 3+ only
'Public Const LVN_ITEMACTIVATE = (LVN_FIRST - 14)
'Public Const LVN_ODSTATECHANGED = (LVN_FIRST - 15)
'
'Public Const LVN_HOTTRACK = (LVN_FIRST - 21)
'
'Public Const LVN_KEYDOWN = (LVN_FIRST - 55)
'Public Const LVN_MARQUEEBEGIN = (LVN_FIRST - 56)        'IE 3+ only
'
''------------------------------------------------------------------------
''ListView Notification Messages (LVN_) (Win95)
'
'Public Const LVN_BEGINLABELEDITA = (LVN_FIRST - 5)
'Public Const LVN_ENDLABELEDITA = (LVN_FIRST - 6)
'
'Public Const LVN_GETDISPINFOA = (LVN_FIRST - 50)
'Public Const LVN_SETDISPINFOA = (LVN_FIRST - 51)
'Public Const LVN_ODFINDITEMA = (LVN_FIRST - 52)       'IE 3+ only
'Public Const LVN_ODFINDITEM = LVN_ODFINDITEMA
'
'Public Const LVN_BEGINLABELEDIT = LVN_BEGINLABELEDITA
'Public Const LVN_ENDLABELEDIT = LVN_ENDLABELEDITA
'Public Const LVN_GETDISPINFO = LVN_GETDISPINFOA
'Public Const LVN_SETDISPINFO = LVN_SETDISPINFOA
'
'Public Const LV_DISPINFOA = NMLVDISPINFOA             'IE 3+ only
'Public Const tagLVDISPINFO = LV_DISPINFO
'Public Const NMLVDISPINFOA = LV_DISPINFOA
'Public Const NMLVDISPINFO = NMLVDISPINFOA

'------------------------------------------------------------------------
'ListView Notification Messages (LVN_) (Unicode)

'Public Const LVN_BEGINLABELEDITW = (LVN_FIRST - 75)
'Public Const LVN_ENDLABELEDITW = (LVN_FIRST - 76)
'
'Public Const LVN_GETDISPINFOW = (LVN_FIRST - 77)
'Public Const LVN_SETDISPINFOW = (LVN_FIRST - 78)
'Public Const LVN_ODFINDITEMW = (LVN_FIRST - 79)         'IE 3+ only
'Public Const LVN_ODFINDITEM = LVN_ODFINDITEMW
'
'Public Const LVN_BEGINLABELEDIT = LVN_BEGINLABELEDITW
'Public Const LVN_ENDLABELEDIT = LVN_ENDLABELEDITW
'Public Const LVN_GETDISPINFO = LVN_GETDISPINFOW
'Public Const LVN_SETDISPINFO = LVN_SETDISPINFOW
'
'Public Const LV_DISPINFOW = NMLVDISPINFOW               'IE 3+ only
'Public Const tagLVDISPINFOW = LV_DISPINFOW
'Public Const NMLVDISPINFOW = LV_DISPINFOW
'Public Const NMLVDISPINFO = NMLVDISPINFOW
