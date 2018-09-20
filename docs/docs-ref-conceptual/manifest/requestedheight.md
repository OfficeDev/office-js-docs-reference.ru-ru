# <a name="requestedheight-element"></a>Элемент RequestedHeight

Задает начальное высоту (в точках) контента надстройки или надстройки почты. 

**Типа надстройки:** Контент, почты

## <a name="syntax"></a>Синтаксис

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a>Содержащиеся в

- [DefaultSettings](defaultsettings.md) (Содержимого надстроек) со значением, которое может быть в диапазоне от 32 до 1000
- [DesktopSettings](desktopsettings.md) и [TabletSettings](tabletsettings.md) (надстройки почты) со значением, которое может быть в диапазоне от 32 до 450
- [ExtensionPoint](extensionpoint.md) (Надстройки контекстной почты) со значением, которое может быть от 140 до 450 пикселей для точки расширения **DetectedEntity** и от 32 до 450 точки расширения **CustomPane**