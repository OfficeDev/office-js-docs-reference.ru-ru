# <a name="extensionpoint-element"></a><span data-ttu-id="efccf-101">Элемент ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="efccf-101">ExtensionPoint element</span></span>

 <span data-ttu-id="efccf-102">Определяет, где доступны функции надстройки в пользовательском интерфейсе Office.</span><span class="sxs-lookup"><span data-stu-id="efccf-102">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="efccf-103">Элемент **ExtensionPoint** является дочерним для элемента [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="efccf-103">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> 

## <a name="attributes"></a><span data-ttu-id="efccf-104">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="efccf-104">Attributes</span></span>

|  <span data-ttu-id="efccf-105">Атрибут</span><span class="sxs-lookup"><span data-stu-id="efccf-105">Attribute</span></span>  |  <span data-ttu-id="efccf-106">Обязательный</span><span class="sxs-lookup"><span data-stu-id="efccf-106">Required</span></span>  |  <span data-ttu-id="efccf-107">Описание</span><span class="sxs-lookup"><span data-stu-id="efccf-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="efccf-108">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="efccf-108">**xsi:type**</span></span>  |  <span data-ttu-id="efccf-109">Да</span><span class="sxs-lookup"><span data-stu-id="efccf-109">Yes</span></span>  | <span data-ttu-id="efccf-110">Тип определяемой точки расширения.</span><span class="sxs-lookup"><span data-stu-id="efccf-110">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="efccf-111">Точки расширения только для Excel</span><span class="sxs-lookup"><span data-stu-id="efccf-111">Extension points for Excel only</span></span>

- <span data-ttu-id="efccf-112">**CustomFunctions** - пользовательские функции, написанной на JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="efccf-112">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="efccf-113">[Пример кода в этой XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml) показано, как использовать элемент **ExtensionPoint** с значение атрибута **CustomFunctions** и дочерние элементы для использования.</span><span class="sxs-lookup"><span data-stu-id="efccf-113">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="efccf-114">Точки расширения для команд надстроек Word, Excel, PowerPoint и OneNote</span><span class="sxs-lookup"><span data-stu-id="efccf-114">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="efccf-115">**PrimaryCommandSurface** — лента в Office.</span><span class="sxs-lookup"><span data-stu-id="efccf-115">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="efccf-116">**ContextMenu** — контекстное меню, которое появляется при нажатии правой кнопкой мыши в интерфейсе Office.</span><span class="sxs-lookup"><span data-stu-id="efccf-116">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="efccf-117">В следующих примерах показано, как использовать элемент **ExtensionPoint** со значениями атрибута **PrimaryCommandSurface** и **ContextMenu**, и какие дочерние элементы использовать с каждым из них.</span><span class="sxs-lookup"><span data-stu-id="efccf-117">The following examples show how to use the  **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="efccf-118">Для элементов, содержащих атрибут ID убедитесь, что предоставление уникального идентификатора.</span><span class="sxs-lookup"><span data-stu-id="efccf-118">For elements that contain an ID attribute, make sure you provide a unique ID.</span></span> <span data-ttu-id="efccf-119">Рекомендуем указать название компании и ваш идентификатор.</span><span class="sxs-lookup"><span data-stu-id="efccf-119">We recommend that you use your company's name along with your ID.</span></span> <span data-ttu-id="efccf-120">Например, используйте следующий формат.</span><span class="sxs-lookup"><span data-stu-id="efccf-120">For example, use the following format.</span></span> <CustomTab id="mycompanyname.mygroupname">

```XML
<ExtensionPoint xsi:type="PrimaryCommandSurface">
          <CustomTab id="Contoso Tab">
          <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
            <!-- <OfficeTab id="TabData"> -->
            <Label resid="residLabel4" />
            <Group id="Group1Id12">
              <Label resid="residLabel4" />
              <Icon>
                <bt:Image size="16" resid="icon1_32x32" />
                <bt:Image size="32" resid="icon1_32x32" />
                <bt:Image size="80" resid="icon1_32x32" />
              </Icon>
              <Tooltip resid="residToolTip" />
              <Control xsi:type="Button" id="Button1Id1">

                  <!-- information about the control -->
              </Control>
              <!-- other controls, as needed -->
            </Group>
          </CustomTab>
        </ExtensionPoint>

      <ExtensionPoint xsi:type="ContextMenu">
        <OfficeMenu id="ContextMenuCell">
          <Control xsi:type="Menu" id="ContextMenu2">
                  <!-- information about the control -->
          </Control>
          <!-- other controls, as needed -->
        </OfficeMenu>
        </ExtensionPoint>
```

#### <a name="child-elements"></a><span data-ttu-id="efccf-121">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="efccf-121">Child elements</span></span>
 
|<span data-ttu-id="efccf-122">**Element**</span><span class="sxs-lookup"><span data-stu-id="efccf-122">**Element**</span></span>|<span data-ttu-id="efccf-123">**Описание**</span><span class="sxs-lookup"><span data-stu-id="efccf-123">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="efccf-124">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="efccf-124">**CustomTab**</span></span>|<span data-ttu-id="efccf-p103">Обязательный, если требуется добавить на ленту настраиваемую вкладку (с помощью элемента **PrimaryCommandSurface**). Если используется элемент **CustomTab**, использовать элемент **OfficeTab** невозможно. Атрибут **id** является обязательным.</span><span class="sxs-lookup"><span data-stu-id="efccf-p103">Required if you want to add a custom tab to the ribbon (using  **PrimaryCommandSurface**). If you use the  **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="efccf-128">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="efccf-128">**OfficeTab**</span></span>|<span data-ttu-id="efccf-p104">Обязательный, если требуется расширить стандартную вкладку ленты Office (с помощью элемента **PrimaryCommandSurface**). Невозможно использовать элементы **OfficeTab** и **CustomTab** одновременно. Дополнительные сведения см. в статье [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="efccf-p104">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the  **OfficeTab** element, you can't use the **CustomTab** element. For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="efccf-132">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="efccf-132">**OfficeMenu**</span></span>|<span data-ttu-id="efccf-p105">Обязательный при добавлении команд надстройки в контекстное меню по умолчанию (с помощью элемента **ContextMenu**). Для атрибута **id** необходимо задать следующее значение: </span><span class="sxs-lookup"><span data-stu-id="efccf-p105">Required if you're adding add-in commands to a default context menu (using  **ContextMenu**). The  **id** attribute must be set to: </span></span><br/> <span data-ttu-id="efccf-p106">- **ContextMenuText** для Excel или Word. Отображает элемент в контекстном меню, когда пользователь щелкает выделенный текст правой кнопкой мыши. </span><span class="sxs-lookup"><span data-stu-id="efccf-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="efccf-p107">- **ContextMenuCell** для Excel. Отображает элемент в контекстном меню, когда пользователь нажимает ячейку электронной таблицы правой кнопкой мыши.</span><span class="sxs-lookup"><span data-stu-id="efccf-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="efccf-139">**Group**</span><span class="sxs-lookup"><span data-stu-id="efccf-139">**Group**</span></span>|<span data-ttu-id="efccf-p108">Группа точек расширения пользовательского интерфейса на вкладке. Группа может включать до шести элементов управления. Атрибут **id** является обязательным. Это строка длиной до 125 символов.</span><span class="sxs-lookup"><span data-stu-id="efccf-p108">A group of user interface extension points on a tab. A group can have up to six controls. The  **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="efccf-143">**Label**</span><span class="sxs-lookup"><span data-stu-id="efccf-143">**Label**</span></span>|<span data-ttu-id="efccf-p109">Обязательный. Метка группы. Для атрибута **resid** необходимо задать значение атрибута **id** элемента **String**. Элемент **String** — это дочерний элемент элемента **ShortStrings**, который является дочерним для элемента **Resources**.</span><span class="sxs-lookup"><span data-stu-id="efccf-p109">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="efccf-148">**Icon**</span><span class="sxs-lookup"><span data-stu-id="efccf-148">**Icon**</span></span>|<span data-ttu-id="efccf-p110">Обязательный. Задает значок группы, который будет использоваться на устройствах с малым форм-фактором либо при отображении слишком большого количества кнопок. Для атрибута **resid** необходимо задать значение атрибута **id** элемента **Image**. Элемент **Image** — это дочерний элемент элемента **Images**, который является дочерним для элемента **Resources**. Атрибут **size** указывает размер изображения в пикселях. Необходимо три размера изображения: 16, 32 и 80. Кроме того, поддерживается пять необязательных размеров: 20, 24, 40, 48 и 64.</span><span class="sxs-lookup"><span data-stu-id="efccf-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="efccf-156">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="efccf-156">**Tooltip**</span></span>|<span data-ttu-id="efccf-p111">Необязательный. Подсказка группы. Для атрибута **resid** необходимо задать значение атрибута **id** элемента **String**. Элемент **String** — это дочерний элемент элемента **LongStrings**, который является дочерним для элемента **Resources**.</span><span class="sxs-lookup"><span data-stu-id="efccf-p111">Optional. The tooltip of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="efccf-161">**Control**</span><span class="sxs-lookup"><span data-stu-id="efccf-161">**Control**</span></span>|<span data-ttu-id="efccf-162">В каждой группе должен быть по крайней мере один элемент управления.</span><span class="sxs-lookup"><span data-stu-id="efccf-162">Each group requires at least one control.</span></span> <span data-ttu-id="efccf-163">Элемент **Control** может относиться к типу **Button** или **Menu**.</span><span class="sxs-lookup"><span data-stu-id="efccf-163">A  **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="efccf-164">С помощью элемента **Menu** можно указать раскрывающийся список элементов управления "Кнопка".</span><span class="sxs-lookup"><span data-stu-id="efccf-164">Use  **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="efccf-165">В настоящее время поддерживаются только кнопки и меню.</span><span class="sxs-lookup"><span data-stu-id="efccf-165">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="efccf-166">Дополнительные сведения см. в разделах [Элементы управления "Кнопка"](control.md#button-control) и [Элементы управления меню](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="efccf-166">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="efccf-167">**Примечание:**  Чтобы сделать, устранение неполадок, рекомендуется, что элемент **управления** и связанных с ними дочерние элементы **ресурсов** будет добавлен по очереди.</span><span class="sxs-lookup"><span data-stu-id="efccf-167">**Note:**  To make troubleshooting easier, we recommend that a  **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="efccf-168">**Script**</span><span class="sxs-lookup"><span data-stu-id="efccf-168">**Script**</span></span>|<span data-ttu-id="efccf-169">Ссылка на файл JavaScript с пользовательским определением функции и кодом регистрации.</span><span class="sxs-lookup"><span data-stu-id="efccf-169">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="efccf-170">Этот элемент не используется в предварительной версии для разработчиков.</span><span class="sxs-lookup"><span data-stu-id="efccf-170">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="efccf-171">Загрузку всех файлов JavaScript выполняет страница HTML.</span><span class="sxs-lookup"><span data-stu-id="efccf-171">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="efccf-172">**Page**</span><span class="sxs-lookup"><span data-stu-id="efccf-172">**Page**</span></span>|<span data-ttu-id="efccf-173">Ссылка на HTML-страницу для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="efccf-173">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="efccf-174">Точки расширения для Outlook</span><span class="sxs-lookup"><span data-stu-id="efccf-174">Extension points for Outlook</span></span>

- <span data-ttu-id="efccf-175">[MessageReadCommandSurface](#messagereadcommandsurface);</span><span class="sxs-lookup"><span data-stu-id="efccf-175">[MessageReadCommandSurface](#messagereadcommandsurface)</span></span> 
- <span data-ttu-id="efccf-176">[MessageComposeCommandSurface](#messagecomposecommandsurface);</span><span class="sxs-lookup"><span data-stu-id="efccf-176">[MessageComposeCommandSurface](#messagecomposecommandsurface)</span></span> 
- <span data-ttu-id="efccf-177">[AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface);</span><span class="sxs-lookup"><span data-stu-id="efccf-177">[AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface)</span></span> 
- <span data-ttu-id="efccf-178">[AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface);</span><span class="sxs-lookup"><span data-stu-id="efccf-178">[AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)</span></span>
- <span data-ttu-id="efccf-179">[Module](#module) (можно использовать только в [DesktopFormFactor](desktopformfactor.md)).</span><span class="sxs-lookup"><span data-stu-id="efccf-179">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="efccf-180">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="efccf-180">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="efccf-181">Events</span><span class="sxs-lookup"><span data-stu-id="efccf-181">Events</span></span>](#events)
- [<span data-ttu-id="efccf-182">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="efccf-182">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="efccf-183">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="efccf-183">MessageReadCommandSurface</span></span>
<span data-ttu-id="efccf-p114">Эта точка расширения помещает кнопки на панель команд для представления чтения почты. В классической версии Outlook эта панель отображается на ленте.</span><span class="sxs-lookup"><span data-stu-id="efccf-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="efccf-186">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="efccf-186">Child elements</span></span>

|  <span data-ttu-id="efccf-187">Элемент</span><span class="sxs-lookup"><span data-stu-id="efccf-187">Element</span></span> |  <span data-ttu-id="efccf-188">Описание</span><span class="sxs-lookup"><span data-stu-id="efccf-188">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="efccf-189">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="efccf-189">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="efccf-190">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="efccf-190">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="efccf-191">CustomTab</span><span class="sxs-lookup"><span data-stu-id="efccf-191">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="efccf-192">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="efccf-192">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="efccf-193">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="efccf-193">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="efccf-194">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="efccf-194">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="efccf-195">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="efccf-195">MessageComposeCommandSurface</span></span>
<span data-ttu-id="efccf-196">Эта точка расширения добавляет кнопки на ленту для надстроек, использующих форму создания сообщения.</span><span class="sxs-lookup"><span data-stu-id="efccf-196">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="efccf-197">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="efccf-197">Child elements</span></span>

|  <span data-ttu-id="efccf-198">Элемент</span><span class="sxs-lookup"><span data-stu-id="efccf-198">Element</span></span> |  <span data-ttu-id="efccf-199">Описание</span><span class="sxs-lookup"><span data-stu-id="efccf-199">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="efccf-200">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="efccf-200">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="efccf-201">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="efccf-201">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="efccf-202">CustomTab</span><span class="sxs-lookup"><span data-stu-id="efccf-202">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="efccf-203">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="efccf-203">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="efccf-204">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="efccf-204">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="efccf-205">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="efccf-205">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="efccf-206">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="efccf-206">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="efccf-207">Эта точка расширения добавляет кнопки на ленту для формы, предназначенной для организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="efccf-207">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="efccf-208">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="efccf-208">Child elements</span></span>

|  <span data-ttu-id="efccf-209">Элемент</span><span class="sxs-lookup"><span data-stu-id="efccf-209">Element</span></span> |  <span data-ttu-id="efccf-210">Описание</span><span class="sxs-lookup"><span data-stu-id="efccf-210">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="efccf-211">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="efccf-211">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="efccf-212">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="efccf-212">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="efccf-213">CustomTab</span><span class="sxs-lookup"><span data-stu-id="efccf-213">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="efccf-214">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="efccf-214">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="efccf-215">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="efccf-215">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="efccf-216">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="efccf-216">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="efccf-217">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="efccf-217">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="efccf-218">Эта точка расширения добавляет кнопки на ленту для формы, предназначенной для участника собрания.</span><span class="sxs-lookup"><span data-stu-id="efccf-218">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="efccf-219">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="efccf-219">Child elements</span></span>

|  <span data-ttu-id="efccf-220">Элемент</span><span class="sxs-lookup"><span data-stu-id="efccf-220">Element</span></span> |  <span data-ttu-id="efccf-221">Описание</span><span class="sxs-lookup"><span data-stu-id="efccf-221">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="efccf-222">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="efccf-222">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="efccf-223">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="efccf-223">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="efccf-224">CustomTab</span><span class="sxs-lookup"><span data-stu-id="efccf-224">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="efccf-225">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="efccf-225">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="efccf-226">Пример элемента OfficeTab</span><span class="sxs-lookup"><span data-stu-id="efccf-226">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="efccf-227">Пример элемента CustomTab</span><span class="sxs-lookup"><span data-stu-id="efccf-227">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="efccf-228">Module</span><span class="sxs-lookup"><span data-stu-id="efccf-228">Module</span></span>

<span data-ttu-id="efccf-229">Эта точка расширения добавляет кнопки на ленту для расширения модуля.</span><span class="sxs-lookup"><span data-stu-id="efccf-229">This extension point puts buttons on the ribbon for the module extension.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="efccf-230">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="efccf-230">Child elements</span></span>

|  <span data-ttu-id="efccf-231">Элемент</span><span class="sxs-lookup"><span data-stu-id="efccf-231">Element</span></span> |  <span data-ttu-id="efccf-232">Описание</span><span class="sxs-lookup"><span data-stu-id="efccf-232">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="efccf-233">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="efccf-233">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="efccf-234">Добавляет команды на вкладку ленты по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="efccf-234">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="efccf-235">CustomTab</span><span class="sxs-lookup"><span data-stu-id="efccf-235">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="efccf-236">Добавляет команды на специальную вкладку ленты.</span><span class="sxs-lookup"><span data-stu-id="efccf-236">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="efccf-237">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="efccf-237">MobileMessageReadCommandSurface</span></span>
<span data-ttu-id="efccf-238">Эта точка расширения помещает кнопки на панель команд для чтения почты в форм-факторе мобильного устройства.</span><span class="sxs-lookup"><span data-stu-id="efccf-238">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="efccf-239">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="efccf-239">Child elements</span></span>

|  <span data-ttu-id="efccf-240">Элемент</span><span class="sxs-lookup"><span data-stu-id="efccf-240">Element</span></span> |  <span data-ttu-id="efccf-241">Описание</span><span class="sxs-lookup"><span data-stu-id="efccf-241">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="efccf-242">Group</span><span class="sxs-lookup"><span data-stu-id="efccf-242">Group</span></span>](group.md) |  <span data-ttu-id="efccf-243">Добавляет группу кнопок на панель команд.</span><span class="sxs-lookup"><span data-stu-id="efccf-243">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="efccf-244">У элементов **ExtensionPoint** этого типа может быть только один дочерний элемент **Group**.</span><span class="sxs-lookup"><span data-stu-id="efccf-244">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="efccf-245">Для атрибута **xsi:type** элементов **Control**, содержащихся в этой точке расширения, должно быть назначено значение `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="efccf-245">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="efccf-246">Пример</span><span class="sxs-lookup"><span data-stu-id="efccf-246">Example</span></span>
```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Group id="mobileGroupID">
    <Label resid="residAppName"/>
      <Control id="mobileButton1" xsi:type="MobileButton">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="events"></a><span data-ttu-id="efccf-247">События</span><span class="sxs-lookup"><span data-stu-id="efccf-247">Events</span></span>

<span data-ttu-id="efccf-248">Эта точка расширения добавляет обработчик для указанного события.</span><span class="sxs-lookup"><span data-stu-id="efccf-248">This extension point adds an event handler for a specified event.</span></span>

> [!NOTE]
> <span data-ttu-id="efccf-249">Данный тип элементов поддерживается только с Outlook в Интернете в Office 365.</span><span class="sxs-lookup"><span data-stu-id="efccf-249">This element type is only supported by Outlook on the web in Office 365.</span></span>

| <span data-ttu-id="efccf-250">Элемент</span><span class="sxs-lookup"><span data-stu-id="efccf-250">Element</span></span> | <span data-ttu-id="efccf-251">Описание</span><span class="sxs-lookup"><span data-stu-id="efccf-251">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="efccf-252">Event</span><span class="sxs-lookup"><span data-stu-id="efccf-252">Event</span></span>](event.md) |  <span data-ttu-id="efccf-253">Задает событие и функцию его обработчика.</span><span class="sxs-lookup"><span data-stu-id="efccf-253">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="efccf-254">Пример события ItemSend</span><span class="sxs-lookup"><span data-stu-id="efccf-254">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events"> 
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
</ExtensionPoint> 
```

### <a name="detectedentity"></a><span data-ttu-id="efccf-255">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="efccf-255">DetectedEntity</span></span>

<span data-ttu-id="efccf-256">Эта точка расширения добавляет активацию контекстной надстройки для указанного типа сущности.</span><span class="sxs-lookup"><span data-stu-id="efccf-256">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="efccf-257">В соответствующем элементе [VersionOverrides](versionoverrides.md) для атрибута `xsi:type` должно быть задано значение `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="efccf-257">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="efccf-258">Данный тип элементов поддерживается только с Outlook в Интернете в Office 365.</span><span class="sxs-lookup"><span data-stu-id="efccf-258">This element type is only supported by Outlook on the web in Office 365.</span></span>

|  <span data-ttu-id="efccf-259">Элемент</span><span class="sxs-lookup"><span data-stu-id="efccf-259">Element</span></span> |  <span data-ttu-id="efccf-260">Описание</span><span class="sxs-lookup"><span data-stu-id="efccf-260">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="efccf-261">Label</span><span class="sxs-lookup"><span data-stu-id="efccf-261">Label</span></span>](#label) |  <span data-ttu-id="efccf-262">Задает метку для надстройки в контекстном окне.</span><span class="sxs-lookup"><span data-stu-id="efccf-262">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="efccf-263">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="efccf-263">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="efccf-264">Задает URL-адрес контекстного окна.</span><span class="sxs-lookup"><span data-stu-id="efccf-264">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="efccf-265">Rule</span><span class="sxs-lookup"><span data-stu-id="efccf-265">Rule</span></span>](rule.md) |  <span data-ttu-id="efccf-266">Задает одно или несколько правил, определяющих, когда активируется надстройка.</span><span class="sxs-lookup"><span data-stu-id="efccf-266">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="efccf-267">Label</span><span class="sxs-lookup"><span data-stu-id="efccf-267">Label</span></span>

<span data-ttu-id="efccf-p115">Обязательный элемент. Метка группы. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="efccf-p115">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="efccf-271">Требования к выделению</span><span class="sxs-lookup"><span data-stu-id="efccf-271">Highlight requirements</span></span>

<span data-ttu-id="efccf-p116">Единственный способ, которым пользователь может активировать контекстную надстройку, — взаимодействие с выделенной сущностью. Разработчики могут указывать, какие сущности выделяются, с помощью атрибута `Highlight` элемента `Rule` для типов правил `ItemHasKnownEntity` и `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="efccf-p116">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="efccf-p117">Однако следует учитывать некоторые ограничения. Они гарантируют, что в соответствующих сообщениях и встречах всегда есть выделенная сущность, с помощью которой пользователь может активировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="efccf-p117">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="efccf-276">Сущности `EmailAddress` и `Url` не поддерживают выделение, поэтому их нельзя использовать для активации надстройки.</span><span class="sxs-lookup"><span data-stu-id="efccf-276">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="efccf-277">Если используется одно правило, то для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="efccf-277">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="efccf-278">Если используется правило `RuleCollection`, совмещенное с другими правилами с помощью оператора `Mode="AND"`, то как минимум в одном из правил для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="efccf-278">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="efccf-279">Если используется правило `RuleCollection`, в котором правила совмещаются с помощью оператора `Mode="OR"`, то в каждом из них для атрибута `Highlight` ДОЛЖНО быть задано значение `all`.</span><span class="sxs-lookup"><span data-stu-id="efccf-279">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="efccf-280">Пример события DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="efccf-280">DetectedEntity event example</span></span>

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint> 
```