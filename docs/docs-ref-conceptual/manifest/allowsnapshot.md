# <a name="allowsnapshot-element"></a><span data-ttu-id="30a10-101">Элемент AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="30a10-101">AllowSnapshot element</span></span>

<span data-ttu-id="30a10-102">Указывает, сохраняется ли моментальный снимок контентной надстройки в документе узла.</span><span class="sxs-lookup"><span data-stu-id="30a10-102">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="30a10-103">**Тип надстройки:** контентные.</span><span class="sxs-lookup"><span data-stu-id="30a10-103">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="30a10-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="30a10-104">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="30a10-105">Содержащиеся в</span><span class="sxs-lookup"><span data-stu-id="30a10-105">Contained in</span></span>

[<span data-ttu-id="30a10-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="30a10-106">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="30a10-107">Замечания</span><span class="sxs-lookup"><span data-stu-id="30a10-107">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="30a10-108">— Это **AllowSnapshot** `true` по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="30a10-108">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="30a10-109">Это делает изображение надстройки отображается для пользователей, откройте документ в версии ведущего приложения, который не поддерживает надстроек Office или содержит Неподвижное изображение элемента надстройки, если ведущему приложению не удается подключиться к на сервере, содержащем надстройки.</span><span class="sxs-lookup"><span data-stu-id="30a10-109">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="30a10-110">Тем не менее это также означает, что потенциально конфиденциальные сведения, отображаемые в надстройку может осуществляться непосредственно из документа, размещения надстройки.</span><span class="sxs-lookup"><span data-stu-id="30a10-110">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

