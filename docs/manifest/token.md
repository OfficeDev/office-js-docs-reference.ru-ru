---
title: Элемент маркера в файле манифеста
description: Указывает маркер или под диктовую карточку, которые можно использовать с шаблонами URL-адресов в манифесте.
ms.date: 11/06/2020
ms.localizationpriority: medium
ms.openlocfilehash: f2064cdd6b9d0a4319ef8c01d056531811c7a8a2
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487453"
---
# <a name="token-element"></a>Элемент Token

Определяет отдельный маркер URL-адреса. Дополнительные сведения об использовании этого элемента см. в дополнительных сведениях о работе с расширенными [переопределениями манифеста](/office/dev/add-ins/develop/extended-overrides).

**Тип надстройки:** надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a>Содержится в

[Tokens](tokens.md)

## <a name="can-contain"></a>Может содержать

|Элемент|Контентная|Почта|Область задач|
|:-----|:-----|:-----|:-----|
|[Override](override.md)|||x|

## <a name="attributes"></a>Атрибуты

|Атрибут|Описание|
|:-----|:-----|
|DefaultValue|Значение по умолчанию для этого маркера, если условие в любом детском `<Override>` элементе не совпадает.|
|Имя|Имя маркера. Это имя определяется пользователем. Тип маркера определяется атрибутом типа.|
|xsi:type|Определяет тип Маркера. Этот атрибут должен быть заданной для одного из:  `"RequirementsToken"`или  `"LocaleToken"`.|

## <a name="example"></a>Пример

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
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
</OfficeApp>
```