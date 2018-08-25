Attribute VB_Name = "MMREG"
'Copyright (c) 1991-1996  Microsoft Corporation

'Module Name:   MMREG.h

'Abstract:      Multimedia Registration

'Revision History:  Copied from MMREG.H and then modified into     Select     Cases

Function Manufacturer(ByVal Mfg_ID As Integer) As String

'Manufacturer IDs

    Select Case Mfg_ID
        Case 1: Manufacturer = "Microsoft Corporation"
        Case 2: Manufacturer = "Creative Labs, Inc"
        Case 21: Manufacturer = "Turtle Beach, Inc."
        Case 22: Manufacturer = "IBM Corporation"
        Case 24: Manufacturer = "Roland"
        Case 25: Manufacturer = "DSP Solutions, Inc."
        Case 26: Manufacturer = "NEC"
        Case 27: Manufacturer = "ATI"
        Case 28: Manufacturer = "Wang Laboratories, Inc"
        Case 30: Manufacturer = "Voyetra"
        Case 33: Manufacturer = "Intel Corporation"
        Case 34: Manufacturer = "Advanced Gravis"
        Case 39: Manufacturer = "Echo Speech Corporation"
        Case 49: Manufacturer = "Canopus Co., Ltd."
        Case 52: Manufacturer = "Aztech Labs, Inc."
        Case 125: Manufacturer = "Ensoniq Corporation"
        Case Else: Manufacturer = "Not Listed"
    End Select

End Function

Function Product(ByVal ProdID As Integer) As String

'Microsoft Product IDs

    Select Case ProdID
        Case 1: Product = "Midi Mapper"
        Case 2: Product = "Wave Mapper"
        Case 3: Product = "Sound Blaster MIDI output port"
        Case 4: Product = "Sound Blaster MIDI input port"
        Case 5: Product = "Sound Blaster internal synth"
        Case 6: Product = "Sound Blaster waveform output"
        Case 7: Product = "Sound Blaster waveform input"
        Case 9: Product = "Ad Lib Compatible synth"
        Case 10: Product = "MPU 401 compatible MIDI output port"
        Case 11: Product = "MPU 401 compatible MIDI input port"
        Case 12: Product = "Joystick adapter"
        Case 13: Product = "PC speaker waveform output"
        Case 14: Product = "MS Audio Board waveform input"
        Case 15: Product = "MS Audio Board waveform output"
        Case 16: Product = "MS Audio Board  Stereo FM synth"
        Case 17: Product = "MS Audio Board Mixer Driver"
        Case 18: Product = "MS OEM Audio Board waveform input"
        Case 19: Product = "MS OEM Audio Board waveform output"
        Case 20: Product = "MS OEM Audio Board Stereo FM Synth"
        Case 21: Product = "MS Audio Board Aux. Port"
        Case 22: Product = "MS OEM Audio Aux Port"
        Case 23: Product = "MS Vanilla driver waveform input"
        Case 24: Product = "MS Vanilla driver wavefrom output"
        Case 25: Product = "MS Vanilla driver MIDI in"
        Case 26: Product = "MS Vanilla driver MIDI  external out"
        Case 27: Product = "MS Vanilla driver MIDI synthesizer"
        Case 28: Product = "MS Vanilla driver aux (line in)"
        Case 29: Product = "MS Vanilla driver aux (mic)"
        Case 30: Product = "MS Vanilla driver aux (CD)"
        Case 31: Product = "MS OEM Audio Board Mixer Driver"
        Case 32: Product = "MS Audio Compression Manager"
        Case 33: Product = "MS ADPCM Codec"
        Case 34: Product = "IMA ADPCM Codec"
        Case 35: Product = "MS Filter"
        Case 36: Product = "GSM 610 codec"
        Case 37: Product = "G.711 codec"
        Case 38: Product = "PCM converter"

 'Microsoft Windows Sound System drivers

        Case 39: Product = "Sound Blaster 16 waveform input"
        Case 40: Product = "Sound Blaster 16  waveform output"
        Case 41: Product = "Sound Blaster 16 midi-in"
        Case 42: Product = "Sound Blaster 16 midi out"
        Case 43: Product = "Sound Blaster 16 FM Synthesis"
        Case 44: Product = "Sound Blaster 16 aux (line in)"
        Case 45: Product = "Sound Blaster 16 aux (CD)"
        Case 46: Product = "Sound Blaster 16 mixer device"
        Case 47: Product = "Sound Blaster Pro waveform input"
        Case 48: Product = "Sound Blaster Pro waveform output"
        Case 49: Product = "Sound Blaster Pro midi in"
        Case 50: Product = "Sound Blaster Pro midi out"
        Case 51: Product = "Sound Blaster Pro FM synthesis"
        Case 52: Product = "Sound Blaster Pro aux (line in )"
        Case 53: Product = "Sound Blaster Pro aux (CD)"
        Case 54: Product = "Sound Blaster Pro mixer"
        Case 55: Product = "Ensoniq SoundscapeVIVO PnP"
        Case 56: Product = "WSS NT wave out"
        Case 57: Product = "WSS NT FM synth"
        Case 58: Product = "WSS NT mixer"
        Case 59: Product = "WSS NT aux"
        Case 60: Product = "Sound Blaster 16 waveform input"
        Case 61: Product = "Sound Blaster 16  waveform output"
        Case 62: Product = "Sound Blaster 16 midi-in"
        Case 63: Product = "Sound Blaster 16 midi out"
        Case 64: Product = "Sound Blaster 16 FM Synthesis"
        Case 65: Product = "Sound Blaster 16 aux (line in)"
        Case 66: Product = "Sound Blaster 16 aux (CD)"
        Case 67: Product = "Sound Blaster 16 mixer device"
        Case 68: Product = "Sound Blaster Pro waveform input"
        Case 69: Product = "Sound Blaster Pro waveform output"
        Case 70: Product = "Sound Blaster Pro midi in"
        Case 71: Product = "Sound Blaster Pro midi out"
        Case 72: Product = "Sound Blaster Pro FM synthesis"
        Case 73: Product = "AudioPCI"
        Case 74: Product = "Sound Blaster Pro aux (CD)"
        Case 75: Product = "Sound Blaster Pro mixer"
        Case 76: Product = "Yamaha OPL2/OPL3 compatible FM synthesis"

'Creative Labs Product IDs

        Case 101: Product = "SB 1.5 Wave Out"
        Case 102: Product = "SB 2.0 Wave Out"
        Case 103: Product = "SB Pro Wave Out"
        Case 104: Product = "SB Pro16 Wave Out"
        Case 201: Product = "SB Midi Out"
        Case 202: Product = "SB Midi In"
        Case 301: Product = "SB FM Synth Mono"
        Case 302: Product = "SB Pro Stereo Synthesizer"
        Case 303: Product = "SB AWE32 Midi"
        Case 401: Product = "SB Pro Aux (CD)"
        Case 402: Product = "SB Pro Aux (Line in)"
        Case 403: Product = "SB Pro Aux (Mic)"
        Case 404: Product = "SB Master Volume"
        Case 405: Product = "SB PC Speaker Volume"
        Case 406: Product = "SB Aux Wave"
        Case 407: Product = "SB Aux Midi"
        Case 408: Product = "SB Pro Mixer"
        Case 409: Product = "SoundBlasterAWE32 Mixer"
        Case Else: Product = "Not Listed"
    End Select
End Function
