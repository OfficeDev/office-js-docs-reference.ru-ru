---
title: Элемент FunctionFile в файле манифеста
description: Указывает исходный файл кода для операций, которые надстройка предоставляет с помощью команд надстройки, которые выполняют функцию JavaScript вместо отображения пользовательского интерфейса.
ms.date: 09/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: ee61192444111a5057ed528bdb743b5ecb8504f5
ms.sourcegitcommit: dc06d41f1d4f0b25bf32d14c1ff377b1e77b7240
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64487261"
---
# <a name="functionfile-element"></a>Элемент FunctionFile

Указывает исходный код файла для операций, которые надстройка предоставляет одним из следующих способов.

* Команды надстройки, которые выполняют функцию JavaScript вместо отображения пользовательского интерфейса.
* Клавиши, которые выполняют функцию JavaScript.

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

Элемент `FunctionFile` является детским элементом [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md). `resid` `FunctionFile` Атрибут элемента может быть не более 32 `id` `Url` `Resources` символов и задарен значению атрибута элемента в элементе, который содержит URL-адрес HTML-файла, который содержит или загружает все функции JavaScript, используемые кнопками командной команды без пользовательского интерфейса, как это определено элементом [Control](control.md).

> [!NOTE]
> Когда надстройка настроена на использование общего времени [работы,](/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime) функции в кодовом файле выполняются в том же времени запуска JavaScript (и совместном общем глобальном пространстве имен), что и JavaScript в области задач надстройки (если таковое имеется).
>
> Элемент `FunctionFile` и связанный с ним файл кода также играют особую роль с помощью настраиваемой клавиши [клавиши,](/office/dev/add-ins/design/keyboard-shortcuts) которые требуют общего времени работы.

Ниже приводится пример элемента `FunctionFile` .

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

JavaScript в HTML-файле `FunctionFile` `Office.initialize` , указанный элементом, должен вызывать и определять именные функции, которые принимают один параметр: `event`. Функции должны использовать API `item.notificationMessages`, чтобы сообщать пользователю о ходе выполнения, успешном завершении или ошибке. Он также должен вызывать метод `event.completed` после выполнения. Имя функций используется в элементе для `FunctionName` кнопок без пользовательского интерфейса.

Ниже приводится пример HTML-файла, определяющий функцию `trackMessage` .

```js
Office.initialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

В следующем коде показано, как реализовать используемую функцию `FunctionName`.

```js
// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

// Your function must be in the global namespace.
function writeText(event) {

    // Implement your custom code here. The following code is a simple example.

    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
```

> [!IMPORTANT]
> Вызов сигналов `event.completed` об успешном обращении с событием. Если функция вызывается несколько раз, например при выборе одной команды надстройки несколько раз, все события автоматически помещаются в очередь. Первое событие запускается автоматически, тогда как остальные ожидают в очереди. При вызове функции `event.completed`выполняется следующий вызов в очереди на эту функцию. Необходимо вызвать `event.completed`; в противном случае функция не будет работать.
