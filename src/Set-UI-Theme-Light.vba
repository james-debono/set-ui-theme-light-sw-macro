'==============================================================================
' UI - Light
'
' Sets the SOLIDWORKS interface brightness theme to Light.
'
' This recolours the SOLIDWORKS chrome - CommandManager, FeatureManager, menus
' and task pane - and deliberately leaves the graphics area alone. Use the
' "Background - *" macros for the graphics area, so the two can be combined
' freely.
'
' This is a system option rather than a document property. It applies to every
' document you open and persists between sessions until it is changed again, and
' nothing is saved into any part, assembly or drawing.
'
' To use, run the macro. A document does not need to be open.
'
' Known quirk: SOLIDWORKS may show a "Script Error" dialog from its own task
' pane when the theme changes. Clicking Yes dismisses it and the theme applies
' correctly. The fault is inside SOLIDWORKS' task pane code, not this macro -
' changing the theme by hand in System Options produces it too.
'
'   Version   0.1.0
'   Date      2026-08-09
'   Author    James Debono
'   Licence   MIT - full text below
'   Source    https://github.com/james-debono/solidworks-themes
'
'------------------------------------------------------------------------------
' CHANGELOG (summary - see CHANGELOG.md for the full history)
'
'   0.1.0   First numbered release.'
'------------------------------------------------------------------------------
' MIT Licence
' SPDX-License-Identifier: MIT
'
' Copyright (c) 2026 James Debono
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'==============================================================================

Option Explicit

Sub main()

    Dim swApp As SldWorks.SldWorks

try_:
    On Error GoTo catch_

    Set swApp = Application.SldWorks

    'Background (interface brightness theme) -> Light
    'NOTE: the enum values are not in the same order as the dropdown -
    'Light=0, Medium=1, Dark=2, MediumLight=3 - so always use the named
    'constant rather than the dropdown position.
    swApp.SetUserPreferenceIntegerValue _
        swUserPreferenceIntegerValue_e.swSystemColorsBackground, _
        swInterfaceBrightnessTheme_e.swInterfaceBrightnessTheme_Light

    GoTo finally_

catch_:
    Debug.Print "UI - Light error: " & Err.Number & " - " & Err.Description
    If Not swApp Is Nothing Then
        swApp.SendMsgToUser2 "UI - Light failed: " & Err.Description, _
            swMessageBoxIcon_e.swMbWarning, swMessageBoxBtn_e.swMbOk
    End If

finally_:

End Sub
