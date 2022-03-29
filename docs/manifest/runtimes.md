---
title: Время запуска в файле манифеста
description: Элемент Runtimes указывает время работы надстройки.
ms.date: 09/28/2021
ms.localizationpriority: medium
ms.openlocfilehash: b0ff18a44753c7c9d8a8f8886e40fb11c0e155d8
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487032"
---
# <a name="runtimes-element"></a>Элемент Runtimes

Указывает время запуска надстройки. Ребенок элемента [`<Host>`](host.md) .

> [!NOTE]
> При запуске Office на Windows, надстройка с элементом в манифесте не обязательно работает в том же элементе управления веб-просмотром`<Runtimes>`, что и в противном случае. Дополнительные сведения о том, как версии Windows и Office, которые обычно используются для управления [веб-просмотром](/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins), см. в Office надстройки. Если условия, описанные там для использования Microsoft Edge с WebView2 (Chromium основе) выполнены, `<Runtimes>` то надстройка использует этот браузер независимо от того, имеет ли он элемент. Однако, если эти условия не выполнены, `<Runtimes>` надстройка с элементом всегда использует Internet Explorer 11 независимо от Windows или Microsoft 365 версии.

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

[Host](host.md)

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
| [Runtime](runtime.md) | Да |  Время запуска надстройки. **Важно**. В настоящее время можно определить только один `<Runtime>` элемент. |

## <a name="see-also"></a>См. также

- [Runtime](runtime.md)
- [Настройка надстройки Office для использования общей среды выполнения JavaScript](/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)
- [Настройка надстройки Outlook для активации на основе событий](/office/dev/add-ins/outlook/autolaunch)
