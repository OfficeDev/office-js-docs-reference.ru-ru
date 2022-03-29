---
title: Элемент Hosts в файле манифеста
description: Указывает Office клиентские приложения, Office надстройка активируется.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 060a98229b90d7f6a8c467af21dfef4c63c922d9
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487480"
---
# <a name="hosts-element"></a>Элемент Hosts

Указывает Office клиентские приложения, Office надстройка активируется. Содержит коллекцию элементов **Host** и их параметров. 

## <a name="as-child-of-versionoverrides-element"></a>Как ребенок элемента VersionOverrides

Сведения в этом разделе применяются только *в том случае* , если элемент **Hosts** является ребенком [VersionOverrides](versionoverrides.md).

Этот элемент переопределяет **элемент Hosts** в базовом манифесте.

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Да   |  Описывает ведущее приложение и его параметры. |
