---
title: Время запуска в файле манифеста
description: Элемент Runtime настраивает надстройку для использования общего времени запуска JavaScript для различных компонентов, например ленты, области задач, пользовательских функций.
ms.date: 03/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1965405724076860b6aad8651cfb8c651f182c53
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487045"
---
# <a name="runtime-element"></a>Элемент runtime

Настраивает надстройку для использования общего времени запуска JavaScript, чтобы все компоненты запускались в одном и том же времени. Ребенок элемента [`<Runtimes>`](runtimes.md) .

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

 - Область задач 1.0
 - Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Связанные с этими наборами требований**:

- [SharedRuntime 1.1](../requirement-sets/common/shared-runtime-requirement-sets.md) (Только при ее использования в надстройке области задач.)

[!include[Runtimes support](../includes/runtimes-note.md)]

## <a name="syntax"></a>Синтаксис

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Содержится в

- [Runtimes](runtimes.md)

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
| [Override](override.md) | Нет | **Outlook**. Указывает расположение URL-адреса файла JavaScript, который Outlook для обработчиков точеч [расширения LaunchEvent](extensionpoint.md#launchevent). **Важно**: в настоящее время можно определить только один `<Override>` элемент и он должен быть типа `javascript`.|

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **resid**  |  Да  | Указывает расположение URL-адреса страницы HTML для надстройки. Символ `resid` может быть не более 32 символов `id` и должен соответствовать атрибуту `Url` элемента элемента `Resources` . |
|  [срок службы](#lifetime-attribute)  |  Нет  | Значение по умолчанию `lifetime` является `short` и не нужно задано. Outlook надстройки активации на основе событий используют только значение`short`. Если вы хотите использовать совместное время работы в Excel надстройки, явно установите значение `long`. |

### <a name="lifetime-attribute"></a>атрибут жизни

Необязательно. Представляет продолжительность разрешенного запуска надстройки.

**Доступные значения**

`short`: По умолчанию. Используется только Outlook надстройки активации на основе событий. После активации надстройки она будет работать в течение максимального количества времени, указанного платформой. В настоящее время это около 5 минут. Это единственное значение, поддерживаемые Outlook.

`long`. Используется только при настройке общего [времени запуска JavaScript](/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime). Надстройка может запускаться в документе, открываемом и запускаемом бесконечно. Например, код области задач будет продолжать работать даже при закрытии области задач пользователем. Это единственное значение, поддерживаемые общим временем запуска.

## <a name="see-also"></a>См. также

- [Runtimes](runtimes.md)
- [Настройка надстройки Office для использования общей среды выполнения JavaScript](/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)
- [Настройка надстройки Outlook для активации на основе событий](/office/dev/add-ins/outlook/autolaunch)
