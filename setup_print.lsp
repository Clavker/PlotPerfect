(defun c:SetupPrint (/ current-layout)
  (setq current-layout (getvar "ctab"))

  (princ "\n==================================================")
  (princ "\nЗадача 2: Установка переменных для печати")
  (princ "\n==================================================")

  (setvar "cmdecho" 0)

  (princ "\n📁 Сохранение чертежа...")
  (command "_.QSAVE")
  (princ " ✅")

  ; Используем команду -SETVAR для принудительной установки
  (princ "\n📄 Установка PDFFRAME = 2...")
  (command "_.-SETVAR" "PDFFRAME" "2")
  (princ " ✅")

  (princ "\n🖼️ Установка IMAGEFRAME = 2...")
  (command "_.-SETVAR" "IMAGEFRAME" "2")
  (princ " ✅")

  (princ "\n🎭 Установка WIPEOUTFRAME = 1...")
  (command "_.-SETVAR" "WIPEOUTFRAME" "1")
  (princ " ✅")

  (princ "\n📐 Установка FRAME = 2...")
  (command "_.-SETVAR" "FRAME" "2")
  (princ " ✅")

  (princ "\n📁 Сохранение изменений...")
  (command "_.QSAVE")
  (princ " ✅")

  (setvar "cmdecho" 1)

  (princ "\n\n==================================================")
  (princ "\n✅ Задача выполнена успешно!")
  (princ "\n==================================================")
  (princ)
)