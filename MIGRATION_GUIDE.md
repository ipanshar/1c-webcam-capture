# Инструкция по переименованию проекта

## ? Выполненные изменения

Проект успешно переименован из **WebCameraCSharp** в **AsterWebCamTo1C**.

### Что было изменено:

1. **Файл проекта**: `WebCameraCSharp.csproj` ? `AsterWebCamTo1C.csproj`
2. **Namespace**: `WebCameraCSharp` ? `AsterWebCamTo1C`
3. **AssemblyName**: `WebCameraCapture` ? `AsterWebCamTo1C`
4. **ProgId**: `WebCameraCapture.WebCameraCapture` ? `AsterWebCamTo1C.WebCameraCapture`
5. **Имя DLL**: `WebCameraCapture.dll` ? `AsterWebCamTo1C.dll`
6. **Документация**: Обновлены README.md, setup.iss, примеры

## ?? Следующие шаги

### 1. Удалить старую регистрацию COM (если была)

Если у вас была зарегистрирована старая версия, удалите её:

```bash
# Запустите от администратора
cd bin\Release\net472
%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe WebCameraCapture.dll /unregister
%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe WebCameraCapture.dll /unregister
```

### 2. Зарегистрировать новый COM компонент

**Способ 1: Использовать готовый скрипт**
```bash
# Запустите от имени администратора
Register_COM.bat
```

**Способ 2: Вручную**
```bash
cd bin\Release\net472
%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe AsterWebCamTo1C.dll /codebase /tlb
%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe AsterWebCamTo1C.dll /codebase /tlb
```

### 3. Обновить код в 1С

**Старый код:**
```1c
КамераOLE = Новый COMОбъект("WebCameraCapture.WebCameraCapture");
КамераOLE.Initialize();
```

**Новый код:**
```1c
КамераOLE = Новый COMОбъект("AsterWebCamTo1C.WebCameraCapture");
КамераOLE.InitializeWithCamera(0);  // Теперь можно указать камеру при инициализации
```

### 4. Проверить работу

```1c
Процедура ТестКамеры()
    КамераOLE = Новый COMОбъект("AsterWebCamTo1C.WebCameraCapture");
    
    Если КамераOLE.InitializeWithCamera(0) Тогда
        Сообщение("? Инициализация успешна!");
        Сообщение("Камер найдено: " + КамераOLE.GetCameraCount());
        
        // Список камер
        СписокКамер = КамераOLE.GetDeviceList();
        Для Индекс = 0 По СписокКамер.UBound() Цикл
            Сообщение("Камера " + Индекс + ": " + СписокКамер[Индекс]);
        КонецЦикла;
        
        КамераOLE.Cleanup();
    Иначе
        Сообщение("? Ошибка: " + КамераOLE.GetErrorMessage());
    КонецЕсли;
КонецПроцедуры
```

## ?? Новые возможности

### 1. Инициализация с выбором камеры
```1c
// Инициализация сразу с нужной камерой
КамераOLE.InitializeWithCamera(1);  // Вторая камера
```

### 2. Изменение размера изображения
```1c
// Захват с изменением размера
ДвоичныеДанные = КамераOLE.CaptureFrameWithSize(640, 480);
```

## ?? Устранение проблем

### Проблема: "Class not registered"
**Решение:**
1. Убедитесь, что скрипт запущен от администратора
2. Проверьте, что файл `AsterWebCamTo1C.dll` существует в `bin\Release\net472\`
3. Перезапустите 1С после регистрации

### Проблема: Старая версия все еще работает
**Решение:**
1. Удалите регистрацию старой версии
2. Перезапустите компьютер (или хотя бы 1С)
3. Зарегистрируйте новую версию

### Проблема: "Specified cast is not valid"
**Решение:**
Проверьте, что вы используете правильный ProgId:
```1c
// Правильно:
КамераOLE = Новый COMОбъект("AsterWebCamTo1C.WebCameraCapture");

// Неправильно (старое название):
КамераOLE = Новый COMОбъект("WebCameraCapture.WebCameraCapture");
```

## ?? Структура файлов

```
D:\visualStudio\Camera\
??? AsterWebCamTo1C.csproj         ? Переименованный файл проекта
??? Register_COM.bat               ? Скрипт регистрации
??? Unregister_COM.bat            ? Скрипт удаления регистрации
??? setup.iss                      ? Обновлен для нового имени
??? README.md                      ? Обновлена документация
??? src\
?   ??? WebCameraCapture.cs       ? Обновлен namespace и ProgId
?   ??? README.md                 ? Обновлена документация
??? examples\
?   ??? example_1c.bsl            ? Обновлены примеры
??? bin\Release\net472\
    ??? AsterWebCamTo1C.dll       ? Новое имя DLL
    ??? DirectShowLib.dll
```

## ? Готово!

Проект успешно переименован и готов к использованию под новым именем **AsterWebCamTo1C**.

Все изменения сохранены, проект собран, осталось только зарегистрировать COM компонент и обновить код в 1С.
