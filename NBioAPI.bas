Attribute VB_Name = "NBioAPI"
' #################################################################################################
'
'   NBioAPI.bas   : Define constants for NBioAPI
'
'   Copyright     : NITGEN Co., Ltd.
'
' #################################################################################################

' -------------------------------------------------------------------------------------------------
'   Error code
' -------------------------------------------------------------------------------------------------

Global Const NBioAPIERROR_NONE = 0
Global Const SecuAPIERROR_NONE = 0


' -------------------------------------------------------------------------------------------------
'   General
' -------------------------------------------------------------------------------------------------

' True / False
Global Const NBioAPI_TRUE = 1
Global Const NBioAPI_FALSE = 0

' -------------------------------------------------------------------------------------------------
'   Device
' -------------------------------------------------------------------------------------------------

' Constant for DeviceID
Global Const NBioAPI_DEVICE_ID_NONE = 0
Global Const NBioAPI_DEVICE_ID_FDP02_0 = 1
Global Const NBioAPI_DEVICE_ID_FDU01_0 = 2
Global Const NBioAPI_DEVICE_ID_AUTO_DETECT = 255

' Constant for Device Name
Global Const NBioAPI_DEVICE_NAME_FDP02 = 1
Global Const NBioAPI_DEVICE_NAME_FDU01 = 2
Global Const NBioAPI_DEVICE_NAME_OSU01 = 3
Global Const NBioAPI_DEVICE_NAME_FDU11 = 4
Global Const NBioAPI_DEVICE_NAME_FSC01 = 5

' -------------------------------------------------------------------------------------------------
'   BSP
' -------------------------------------------------------------------------------------------------

' Constant for Security Level
Global Const NBioAPI_FIR_SECURITY_LEVEL_LOWEST = 1
Global Const NBioAPI_FIR_SECURITY_LEVEL_LOWER = 2
Global Const NBioAPI_FIR_SECURITY_LEVEL_LOW = 3
Global Const NBioAPI_FIR_SECURITY_LEVEL_BELOW_NORMAL = 4
Global Const NBioAPI_FIR_SECURITY_LEVEL_NORMAL = 5
Global Const NBioAPI_FIR_SECURITY_LEVEL_ABOVE_NORMAL = 6
Global Const NBioAPI_FIR_SECURITY_LEVEL_HIGH = 7
Global Const NBioAPI_FIR_SECURITY_LEVEL_HIGHER = 8
Global Const NBioAPI_FIR_SECURITY_LEVEL_HIGHEST = 9

' Purpose for FIR
Global Const NBioAPI_FIR_PURPOSE_VERIFY = 1
Global Const NBioAPI_FIR_PURPOSE_IDENTIFY = 2
Global Const NBioAPI_FIR_PURPOSE_ENROLL = 3
Global Const NBioAPI_FIR_PURPOSE_ENROLL_FOR_VERIFICATION_ONLY = 4
Global Const NBioAPI_FIR_PURPOSE_ENROLL_FOR_IDENTIFICATION_ONLY = 5
Global Const NBioAPI_FIR_PURPOSE_AUDIT = 6
Global Const NBioAPI_FIR_PURPOSE_UPDATE = 10

' Finger ID
Global Const NBioAPI_FINGER_ID_UNKNOWN = 0
Global Const NBioAPI_FINGER_ID_RIGHT_THUMB = 1
Global Const NBioAPI_FINGER_ID_RIGHT_INDEX = 2
Global Const NBioAPI_FINGER_ID_RIGHT_MIDDLE = 3
Global Const NBioAPI_FINGER_ID_RIGHT_RING = 4
Global Const NBioAPI_FINGER_ID_RIGHT_LITTLE = 5
Global Const NBioAPI_FINGER_ID_LEFT_THUMB = 6
Global Const NBioAPI_FINGER_ID_LEFT_INDEX = 7
Global Const NBioAPI_FINGER_ID_LEFT_MIDDLE = 8
Global Const NBioAPI_FINGER_ID_LEFT_RING = 9
Global Const NBioAPI_FINGER_ID_LEFT_LITTLE = 10

' Window Style
Global Const NBioAPI_WINDOW_STYLE_POPUP = 0
Global Const NBioAPI_WINDOW_STYLE_INVISIBLE = 1     'only for NBioAPI_Capture()
Global Const NBioAPI_WINDOW_STYLE_CONTINUOUS = 2

Global Const NBioAPI_WINDOW_STYLE_NO_FPIMG = 65536
Global Const NBioAPI_WINDOW_STYLE_TOPMOST = 131072  ' currently not used (after v2.3)
Global Const NBioAPI_WINDOW_STYLE_NO_WELCOME = 262144
Global Const NBioAPI_WINDOW_STYLE_NO_TOPMOST = 524288

' -------------------------------------------------------------------------------------------------
'   Export Data
' -------------------------------------------------------------------------------------------------
Global Const MINCONV_TYPE_FDP = 0
Global Const MINCONV_TYPE_FDU = 1
Global Const MINCONV_TYPE_FDA = 2
Global Const MINCONV_TYPE_OLD_FDA = 3
Global Const MINCONV_TYPE_FDAC = 4
Global Const MINCONV_TYPE_FIM10_HV = 5
Global Const MINCONV_TYPE_FIM10_LV = 6
Global Const MINCONV_TYPE_FIM01_HV = 7
Global Const MINCONV_TYPE_FIM01_HD = 8
Global Const MINCONV_TYPE_FELICA = 9
' -------------------------------------------------------------------------------------------------
'   Export Image
' -------------------------------------------------------------------------------------------------

' Constant for FP Image
Global Const NBioAPI_IMG_TYPE_RAW = 1
Global Const NBioAPI_IMG_TYPE_BMP = 2
Global Const NBioAPI_IMG_TYPE_JPG = 3

