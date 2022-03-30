---
title: Элемент Override в файле манифеста
description: Элемент Переопределения позволяет указать значение параметра в зависимости от заданного условия.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: f6a13d981f249f2095136bed8e1adc5c02332d8f
ms.sourcegitcommit: fe8d1195ab75ae24f88ac54b1d40e0a89dae2d89
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/28/2022
ms.locfileid: "64503493"
---
# <a name="override-element"></a>Элемент Override

Предоставляет способ переопределения значения параметра манифеста в зависимости от указанного условия. Существует три типа условий:

- Локальный Office, `LocaleToken`который отличается от по умолчанию, называется **LocaleTokenOverride**.
- Шаблон поддержки набора требований, который отличается от `RequirementToken` шаблона по умолчанию, **называется RequirementTokenOverride**.
- Источник отличается от по умолчанию `Runtime`, называемый **RuntimeOverride**.

Элемент **Переопределения** , который находится внутри элемента **Runtime** , должен иметь тип **RuntimeOverride**.

Нет атрибута `overrideType` элемента **Переопределения** . Разница определяется родительским элементом и типом родительского элемента. Элемент **Переопределения** , который находится внутри элемента **Token** , `xsi:type` `RequirementToken`который является , должен иметь тип **RequirementTokenOverride**. Элемент **Переопределения** в любом другом родительском элементе или **элементе override** `LocaleToken`типа должен иметь тип **LocaleTokenOverride**. Дополнительные сведения об использовании этого элемента, когда он является ребенком элемента **Token** , см. в этой ссылке Работа с расширенными переопределениями [манифеста](/office/dev/add-ins/develop/extended-overrides).

Каждый тип описан в отдельных разделах позднее в этой статье.

## <a name="override-element-for-localetoken"></a>Элемент Переопределения для `LocaleToken`

Элемент **Переопределения** выражает условный и может быть прочитано как "Если ... затем ..." заявление. Если элемент **Переопределения** имеет тип **LocaleTokenOverride**, `Locale` атрибут — это условие, `Value` а атрибут — следовательно. Например, ниже приводится следующий текст: "Если параметр Office является fr-fr, то имя отображения — "Lecteur vidéo".

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

### <a name="syntax"></a>Синтаксис

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a>Содержится в

|Элемент|
|:-----|
|[CitationText](citationtext.md)|
|[Описание](description.md)|
|[DictionaryName](dictionaryname.md)|
|[DictionaryHomePage](dictionaryhomepage.md)|
|[DisplayName](displayname.md)|
|[HighResolutionIconUrl](highresolutioniconurl.md)|
|[IconUrl](iconurl.md)|
|[QueryUri](queryuri.md)|
|[SourceLocation](sourcelocation.md)|
|[SupportUrl](supporturl.md)|
|[Маркер](token.md)|

### <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|Языковой стандарт|string|Обязательный|Задает имя языка и региональных параметров для языкового стандарта этого переопределения в формате языковых тегов BCP 47, например `"en-US"`.|
|Значение|string|Обязательный|Задает значение параметра, представленное для указанного языкового стандарта.|

### <a name="examples"></a>Примеры

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

```xml
<bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
    <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
</bt:Image>
```

```xml
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
```

### <a name="see-also"></a>См. также

- [Локализация надстроек для Office](/office/dev/add-ins/develop/localization)
- [Сочетания клавиш](/office/dev/add-ins/design/keyboard-shortcuts)

## <a name="override-element-for-requirementtoken"></a>Элемент Переопределения для `RequirementToken`

Элемент **Переопределения** выражает условный и может быть прочитано как "Если ... затем ..." заявление. Если элемент **Переопределения** имеет тип **RequirementTokenOverride**, то **элемент требования к** ребенку выражает условие, `Value` а атрибут — следовательно. Например, первое переопределения в следующем: "Если текущая платформа поддерживает версию FeatureOne 1.7, используйте строку "oldAddinVersion"  `${token.requirements}` `<ExtendedOverrides>` вместо маркера в URL-адресе бабушки и дедушки (вместо строки по умолчанию "обновление")".

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Tokens>
        <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
            <Override Value="oldAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.7" />
                    </Sets>
                </Requirements>
            </Override>
            <Override Value="currentAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.8" />
                    </Sets>
                    <Methods>
                        <Method Name="MethodThree" />
                    </Methods>
                </Requirements>
            </Override>
        </Token>
    </Tokens>
</ExtendedOverrides>
```

**Тип надстройки:** надстройки области задач

### <a name="syntax"></a>Синтаксис

```XML
<Override Value="string" />
```

### <a name="contained-in"></a>Содержится в

|Элемент|
|:-----|
|[Маркер](token.md)|

### <a name="must-contain"></a>Должен содержать

|Элемент|Контентная|Почта|Область задач|
|:-----|:-----|:-----|:-----|
|[Requirements](requirements.md)|||x|

### <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|Значение|string|Обязательный|Значение маркера дедушек и дедушек при условии удовлетворены.|

### <a name="example"></a>Пример

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
        <Override Value="very-old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.5" />
                    <Set Name="FeatureTwo" MinVersion="1.1" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.7" />
                    <Set Name="FeatureTwo" MinVersion="1.2" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="current">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.8" />
                    <Set Name="FeatureTwo" MinVersion="1.3" />
                </Sets>
                <Methods>
                    <Method Name="MethodThree" />
                </Methods>
            </Requirements>
        </Override>
    </Token>
</ExtendedOverrides>
```

### <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Укажите Office какие версии и платформы могут принимать вашу надстройка](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#specify-which-office-versions-and-platforms-can-host-your-add-in)
- [Сочетания клавиш](/office/dev/add-ins/design/keyboard-shortcuts)

## <a name="override-element-for-runtime"></a>Элемент Переопределения для `Runtime`

> [!IMPORTANT]
> Поддержка этого элемента была представлена в [наборе требований к почтовым ящикам 1.10](../requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10.md) с функцией активации на [основе событий](/office/dev/add-ins/outlook/autolaunch). См [клиенты и платформы](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

Элемент **Переопределения** выражает условный и может быть прочитано как "Если ... затем ..." заявление. Если элемент **Переопределения** имеет тип **RuntimeOverride**, `type` атрибут — это условие, `resid` а атрибут — следовательно. Например, ниже приводится следующее: "Если тип является "javascript", `resid` то это "JSRuntime.Url". Outlook Desktop требует этот элемент для [обработчиков точек расширения LaunchEvent](/office/dev/add-ins/reference/manifest/extensionpoint#launchevent).

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

**Тип надстройки:** почтовая

### <a name="syntax"></a>Синтаксис

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### <a name="contained-in"></a>Содержится в

- [Runtime](runtime.md)

### <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|**type**|string|Да|Указывает язык для этого переопределения. В настоящее время `"javascript"` это единственный поддерживаемый вариант.|
|**resid**|string|Да|Указывает расположение URL-адреса файла JavaScript, который должен переопределять расположение URL-адреса HTML по умолчанию, определенного в родительском элементе [Runtime](runtime.md)`resid`. Символ `resid` может быть не более 32 символов `id` и должен соответствовать атрибуту `Url` элемента элемента `Resources` .|

### <a name="examples"></a>Примеры

```xml
<!-- Event-based activation happens in a lightweight runtime.-->
<Runtimes>
  <!-- HTML file including reference to or inline JavaScript event handlers.
  This is used by Outlook on the web and Outlook on the new Mac UI preview. -->
  <Runtime resid="WebViewRuntime.Url">
    <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
    <Override type="javascript" resid="JSRuntime.Url"/>
  </Runtime>
</Runtimes>
```

### <a name="see-also"></a>См. также

- [Runtime](runtime.md)
- [Настройка надстройки Outlook для активации на основе событий](/office/dev/add-ins/outlook/autolaunch)
