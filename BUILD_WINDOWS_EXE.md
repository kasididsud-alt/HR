# วิธีสร้างไฟล์ Windows EXE

ไฟล์ exe ต้องสร้างบน Windows เพราะ PyInstaller ไม่รองรับการ cross-build จาก macOS ไปเป็น Windows exe โดยตรง

## ขั้นตอน

1. ติดตั้ง Python 3.10 ขึ้นไปบน Windows
2. เปิดโฟลเดอร์โปรเจกต์นี้บน Windows
3. ดับเบิลคลิก `build_windows_exe.bat`
4. รอจนเสร็จ จะได้ไฟล์ที่ `dist\SalaryCalc_Portable.zip`
5. แตก zip แล้วเปิด `SalaryCalc.exe`

ไฟล์ build จะใช้ icon เดิมจาก `assets\salary_calc.ico`

ถ้าต้องการรันผ่าน PowerShell เอง ใช้คำสั่งนี้:

```powershell
.\build_onefile.ps1
```
