# UserForms-Class-Spliter VBA Project

**Русский** | [English](README.md) | [UserForms-Class-ALL](https://github.com/vbatools/UserForms-Class-ALL/blob/main/README_RUS.md)

![Демонстрация проекта](User_Forms.gif)

## Описание

Проект представляет собой библиотеку для создания сплиттеров (разделителей) между элементами управления на UserForm в Excel. Сплиттеры позволяют пользователям динамически изменять размеры элементов управления на форме.

## Структура проекта

- `clsSpliter.cls` - основной класс-менеджер сплиттеров
- `clsSpliterItem.cls` - класс отдельного элемента сплиттера
- `frmSpliter.frm` - форма с примером использования сплиттеров
- `modShowForms.bas` - модуль для отображения формы

## Основные компоненты

### clsSpliter
- Управляет коллекцией сплиттеров
- Предоставляет методы для добавления, удаления и изменения свойств сплиттеров
- Поддерживает массовые операции над всеми сплиттерами

### clsSpliterItem
- Представляет отдельный элемент сплиттера
- Управляет двумя элементами управления, которые разделяет
- Обрабатывает перемещение сплиттера изменение размеров элементов

### frmSpliter
- Форма с примером использования сплиттеров
- Демонстрирует функциональность библиотеки
- Содержит элементы управления для тестирования сплиттеров

## Использование

### Создание сплиттера
```vba
Dim spliter As New clsSpliter
Call spliter.Initialize(UserForm1) ' или другой родительский объект
Call spliter.AddItem(Control1, Control2) ' создание сплиттера между двумя контролами
```

### Управление сплиттерами
```vba
' Установка цвета текста для всех сплиттеров
Call spliter.setForeColorAll(vbBlue)

' Установка заголовка для всех сплиттеров
Call spliter.setCaptionAll("Разделитель")

' Управление видимостью
Call spliter.setVisibleAll(True)

' Управление доступностью
Call spliter.setEnabledAll(True)

' Удаление всех сплиттеров
Call spliter.RemoveAll()
```

### Управление отдельным сплиттером
```vba
' Получение сплиттера по индексу
Dim item As clsSpliterItem
Set item = spliter.item(1)

' Изменение свойств отдельного сплиттера
item.Caption = "Новый заголовок"
item.ForeColor = vbRed
item.Visible = True
item.Enabled = True

' Или через методы clsSpliter
Call spliter.ByItemSetCaption(1, "Заголовок")
Call spliter.ByItemSetForeColor(1, vbRed)
Call spliter.ByItemSetVisible(1, True)
Call spliter.ByItemSetEnabled(1, True)
```

## Тестирование

Для тестирования функциональности проекта можно использовать форму frmSpliter:
1. Запустить макрос `showForm` из модуля `modShowForms`
2. Использовать элементы управления на форме для проверки работы сплиттеров

## Совместимость

Проект разработан для использования в Microsoft Excel с VBA. Совместим с Excel 2007 и более поздними версиями.

## Лицензия

Apache License
