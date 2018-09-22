# <a name="desktopformfactor-element"></a>Элемент DesktopFormFactor

Указывает параметры для надстройки классического форм-фактора. Классический форм-фактор включает Office для Windows, Office для Mac и Office Online. Он содержит все сведения о надстройке для классического форм-фактора, кроме узла **Resources**.

В каждом определении DesktopFormFactor есть элемент **FunctionFile**, а также один или несколько элементов **ExtensionPoint**. Дополнительные сведения см. в статьях [Элемент FunctionFile](functionfile.md) и [Элемент ExtensionPoint](extensionpoint.md).

> [!IMPORTANT]
> Элемент SupportsSharedFolders доступна только в Outlook надстроек Preview требований в Exchange Online.
> Надстройки, использующих этот элемент не разрешается в магазин Office или централизованного развертывания.

## <a name="child-elements"></a>Дочерние элементы

| Элемент                               | Обязательный | Описание  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | Да      | Определяет, где предоставляются функции надстройки. |
| [FunctionFile](functionfile.md)       | Да      | URL-адрес файла, который содержит функции JavaScript.|
| [GetStarted](getstarted.md)           | Нет       | Определяет выноску, которая отображается при установке надстройки в ведущих приложениях Word, Excel и PowerPoint. |
| SupportsSharedFolders                 | Нет       | Определяет доступна в сценариях делегат надстройки Outlook и имеет значение *false* по умолчанию. Предварительный просмотр наборы требований.|

## <a name="desktopformfactor-example"></a>Пример DesktopFormFactor

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
