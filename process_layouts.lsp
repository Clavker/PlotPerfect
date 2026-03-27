(defun c:PPGO (/ layouts original-layout success-list error-list back-point front-point result-file)
  (setq layouts (layoutlist)
        original-layout (getvar "ctab")
        back-point '(1.06 3.51)
        front-point '(8 5)
        success-list '()
        error-list '()
        result-file (strcat (getenv "TEMP") "\\acad_results.txt"))

  (princ "\n==================================================")
  (princ "\nЗадача 1: Выставить рамки для печати")
  (princ "\n==================================================")
  (princ (strcat "\nНайдено листов для обработки: " (itoa (length layouts))))

  (foreach layout layouts
    (if (/= layout "Model")
      (progn
        (princ (strcat "\n\n--- Обработка листа: " layout " ---"))
        (setvar "ctab" layout)

        ; Обработка рамки чертежа (на передний план)
        (setq success nil)
        (repeat 3
          (if (not success)
            (progn
              (command "_.delay" "100")
              (if (ssget "_C" front-point front-point)
                (progn
                  (command "_.DRAWORDER" "_P" "" "_Front")
                  (while (> (getvar "cmdactive") 0) (command ""))
                  (setq success t)
                  (princ "\n  ✓ Рамка чертежа выведена на передний план")
                )
              )
            )
          )
        )

        ; Обработка номера листа (на задний план)
        (setq success2 nil)
        (repeat 3
          (if (not success2)
            (progn
              (command "_.delay" "100")
              (if (ssget "_C" back-point back-point)
                (progn
                  (command "_.DRAWORDER" "_P" "" "_Back")
                  (while (> (getvar "cmdactive") 0) (command ""))
                  (setq success2 t)
                  (princ "\n  ✓ Номер листа отправлен на задний план")
                )
              )
            )
          )
        )

        ; Формируем результат для листа
        (if (and success success2)
          (progn
            (setq success-list (cons layout success-list))
            (princ "\n  ✓ Лист обработан успешно")
          )
          (progn
            (setq error-list (cons layout error-list))
            (princ "\n  ✗ Ошибка при обработке листа")
            (if (not success) (princ "\n    - Рамка чертежа не найдена"))
            (if (not success2) (princ "\n    - Номер листа не найден"))
          )
        )
      )
    )
  )

  (setvar "ctab" original-layout)

  ; Сохраняем результаты в файл
  (setq file (open result-file "w"))
  (write-line "==================================================" file)
  (write-line "=== ОТЧЕТ ПО ОБРАБОТКЕ ЛИСТОВ ===" file)
  (write-line "==================================================" file)
  (write-line (strcat "Дата: " (menucmd "M=$(edtime,$(getvar,date),DD.MM.YYYY HH:MM:SS)")) file)
  (write-line (strcat "Успешно обработано: " (itoa (length success-list)) " листов") file)
  (write-line (strcat "Ошибок: " (itoa (length error-list)) " листов") file)

  (if success-list
    (progn
      (write-line "" file)
      (write-line "=== УСПЕШНЫЕ ЛИСТЫ ===" file)
      (foreach layout (reverse success-list)
        (write-line (strcat "  • " layout) file)
      )
    )
  )

  (if error-list
    (progn
      (write-line "" file)
      (write-line "=== ЛИСТЫ С ОШИБКАМИ ===" file)
      (foreach layout (reverse error-list)
        (write-line (strcat "  • " layout) file)
      )
    )
  )

  (write-line "==================================================" file)
  (close file)

  ; Итоговый отчет в консоль AutoCAD
  (princ "\n\n==================================================")
  (princ "\nИТОГОВЫЙ ОТЧЕТ")
  (princ "\n==================================================")
  (princ (strcat "\n✅ Успешно обработано: " (itoa (length success-list)) " листов"))
  (princ (strcat "\n❌ Ошибок: " (itoa (length error-list)) " листов"))

  (if success-list
    (progn
      (princ "\n\n✅ Успешные листы:")
      (foreach layout (reverse success-list)
        (princ (strcat "\n  • " layout))
      )
    )
  )

  (if error-list
    (progn
      (princ "\n\n❌ Ошибки:")
      (foreach layout (reverse error-list)
        (princ (strcat "\n  • " layout))
      )
    )
  )

  (princ "\n\n==================================================")
  (princ "\n✅ Обработка завершена!")
  (princ)
)