# <a name="permissions-element"></a><span data-ttu-id="f37a3-101">Элемент Permissions</span><span class="sxs-lookup"><span data-stu-id="f37a3-101">Permissions element</span></span>

<span data-ttu-id="f37a3-102">Указывает уровень доступа к API для надстройки Office. Запрашивая разрешения, руководствуйтесь принципом минимальных привилегий.</span><span class="sxs-lookup"><span data-stu-id="f37a3-102">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="f37a3-103">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="f37a3-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f37a3-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="f37a3-104">Syntax</span></span>

<span data-ttu-id="f37a3-105">Для надстроек области задач и контентных надстроек:</span><span class="sxs-lookup"><span data-stu-id="f37a3-105">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="f37a3-106">Для надстроек почты</span><span class="sxs-lookup"><span data-stu-id="f37a3-106">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="f37a3-107">Содержащиеся в</span><span class="sxs-lookup"><span data-stu-id="f37a3-107">Contained in</span></span>

[<span data-ttu-id="f37a3-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="f37a3-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="f37a3-109">Замечания</span><span class="sxs-lookup"><span data-stu-id="f37a3-109">Remarks</span></span>

<span data-ttu-id="f37a3-110">Подробные сведения см. в статьях [Запрашивание разрешений на использование API в надстройках области задач и контентных надстройках](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) и [Общие сведения о разрешениях для надстроек Outlook](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="f37a3-110">For more detail, see [Requesting permissions for API use in content and task pane add-ins](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
