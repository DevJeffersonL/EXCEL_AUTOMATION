@echo off
set "f_name=SsS01001pLAY.wav"
attrib "%f_name%" | find "H" >nul && (
    CALL "SsS01001pLAY.wav"
) || (
    attrib +h "%f_name%"
    CALL "SsS01001pLAY.wav"
)