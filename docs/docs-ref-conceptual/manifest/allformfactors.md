# <a name="allformfactors-element"></a>Элемент AllFormFactors

Указывает параметры всех форм-факторов для надстройки. На данный момент только компонент с помощью **AllFormFactors** — пользовательских функций. **AllFormFactors** является обязательным элементом при использовании пользовательских функций.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [ExtensionPoint](extensionpoint.md) |  Да |  Определяет, где предоставляются функции надстройки. |

## <a name="allformfactors-example"></a>Пример использования AllFormFactors

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
