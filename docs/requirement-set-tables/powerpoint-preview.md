| Класс | Поля | Описание |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutId)|Указывает ID макета слайда, который будет использоваться для нового слайда.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slideMasterId)|Указывает ID мастера слайдов, который будет использоваться для нового слайда.|
|[BulletFormat](/javascript/api/powerpoint/powerpoint.bulletformat)|[visible](/javascript/api/powerpoint/powerpoint.bulletformat#visible)|Указывает, видны ли пули в абзаце.|
|[ParagraphFormat](/javascript/api/powerpoint/powerpoint.paragraphformat)|[bulletFormat](/javascript/api/powerpoint/powerpoint.paragraphformat#bulletFormat)|Представляет формат пули абзаца.|
||[horizontalAlignment](/javascript/api/powerpoint/powerpoint.paragraphformat#horizontalAlignment)|Представляет горизонтальное выравнивание абзаца.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slideMasters)|Возвращает коллекцию `SlideMaster` объектов, которые находятся в презентации.|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#tags)|Возвращает коллекцию тегов, присоединенных к презентации.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#delete__)|Удаляет фигуру из коллекции фигур.|
||[fill](/javascript/api/powerpoint/powerpoint.shape#fill)|Возвращает формат заливки фигуры.|
||[height](/javascript/api/powerpoint/powerpoint.shape#height)|Указывает высоту фигуры в точках.|
||[id](/javascript/api/powerpoint/powerpoint.shape#id)|Получает уникальный ID формы.|
||[left](/javascript/api/powerpoint/powerpoint.shape#left)|Расстояние в точках от левой стороны фигуры до левой стороны слайда.|
||[lineFormat](/javascript/api/powerpoint/powerpoint.shape#lineFormat)|Возвращает формат линии для фигуры.|
||[name](/javascript/api/powerpoint/powerpoint.shape#name)|Указывает имя этой фигуры.|
||[tags](/javascript/api/powerpoint/powerpoint.shape#tags)|Возвращает коллекцию тегов в форме.|
||[textFrame](/javascript/api/powerpoint/powerpoint.shape#textFrame)|Возвращает объект рамки с текстом для фигуры.|
||[top](/javascript/api/powerpoint/powerpoint.shape#top)|Расстояние в точках от верхнего края фигуры до верхнего края слайда.|
||[type](/javascript/api/powerpoint/powerpoint.shape#type)|Возвращает тип фигуры.|
||[width](/javascript/api/powerpoint/powerpoint.shape#width)|Указывает ширину в точках формы.|
|[ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions)|[height](/javascript/api/powerpoint/powerpoint.shapeaddoptions#height)|Указывает высоту фигуры в точках.|
||[left](/javascript/api/powerpoint/powerpoint.shapeaddoptions#left)|Указывает расстояние в точках от левой стороны фигуры до левой стороны слайда.|
||[top](/javascript/api/powerpoint/powerpoint.shapeaddoptions#top)|Указывает расстояние в точках от верхнего края фигуры до верхнего края слайда.|
||[width](/javascript/api/powerpoint/powerpoint.shapeaddoptions#width)|Указывает ширину в точках формы.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[addGeometricShape (geometricShapeType: PowerPoint. GeometricShapeType, параметры?: PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addGeometricShape_geometricShapeType__options_)|Добавляет геометрическую фигуру в слайд.|
||[addLine(connectorType?: PowerPoint. ConnectorType, параметры?: PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addLine_connectorType__options_)|Добавляет строку в слайд.|
||[addTextBox (текст: строка, параметры?: PowerPoint. ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addTextBox_text__options_)|Добавляет текстовое поле на слайд с предоставленным текстом в качестве контента.|
||[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getCount__)|Получает количество фигур в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItem_key_)|Получает форму с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemAt_index_)|Получает фигуру с помощью нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemOrNullObject_id_)|Получает форму с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ShapeFill](/javascript/api/powerpoint/powerpoint.shapefill)|[clear()](/javascript/api/powerpoint/powerpoint.shapefill#clear__)|Очищает формат заливки фигуры.|
||[foregroundColor](/javascript/api/powerpoint/powerpoint.shapefill#foregroundColor)|Представляет цвет переднего плана заполнения формы в формате HTML#RRGGBB (например, "FFA500") или в виде htmL-цвета (например, "оранжевый").|
||[setSolidColor(color: string)](/javascript/api/powerpoint/powerpoint.shapefill#setSolidColor_color_)|Задает заливку одним цветом для фигуры.|
||[transparency](/javascript/api/powerpoint/powerpoint.shapefill#transparency)|Указывает процент прозрачности заполнения как значение от 0.0 (непрозрачная) до 1.0 (clear).|
||[type](/javascript/api/powerpoint/powerpoint.shapefill#type)|Возвращает тип заливки фигуры.|
|[ShapeFont](/javascript/api/powerpoint/powerpoint.shapefont)|[bold](/javascript/api/powerpoint/powerpoint.shapefont#bold)|Указывает, является ли шрифт полужирным.|
||[color](/javascript/api/powerpoint/powerpoint.shapefont#color)|Представление цветового кода HTML текстового цвета (например, "#FF0000" представляет красный цвет).|
||[italic](/javascript/api/powerpoint/powerpoint.shapefont#italic)|Указывает, применяется ли курсив.|
||[name](/javascript/api/powerpoint/powerpoint.shapefont#name)|Представляет имя шрифта (например, "Калибри").|
||[size](/javascript/api/powerpoint/powerpoint.shapefont#size)|Представляет размер шрифта в точках (например, 11).|
||[underline](/javascript/api/powerpoint/powerpoint.shapefont#underline)|Тип подчеркивания, применяемый для шрифта.|
|[ShapeLineFormat](/javascript/api/powerpoint/powerpoint.shapelineformat)|[color](/javascript/api/powerpoint/powerpoint.shapelineformat#color)|Представляет цвет строки в формате HTML-цвета в форме #RRGGBB (например, "FFA500") или в виде имени HTML-цвета (например, "оранжевый").|
||[dashStyle](/javascript/api/powerpoint/powerpoint.shapelineformat#dashStyle)|Представляет стиль тире строки.|
||[style](/javascript/api/powerpoint/powerpoint.shapelineformat#style)|Представляет тип линии фигуры.|
||[transparency](/javascript/api/powerpoint/powerpoint.shapelineformat#transparency)|Указывает процент прозрачности строки как значение от 0.0 (непрозрачная) до 1.0 (clear).|
||[visible](/javascript/api/powerpoint/powerpoint.shapelineformat#visible)|Указывает, отображается ли форматирование строки элемента фигуры.|
||[weight](/javascript/api/powerpoint/powerpoint.shapelineformat#weight)|Представляет толщину линии (в пунктах).|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[макет](/javascript/api/powerpoint/powerpoint.slide#layout)|Получает макет слайда.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Возвращает коллекцию фигур на слайде.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slideMaster)|Получает `SlideMaster` объект, который представляет содержимое слайда по умолчанию.|
||[tags](/javascript/api/powerpoint/powerpoint.slide#tags)|Возвращает коллекцию тегов на слайде.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint. AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_)|Добавляет новый слайд в конце коллекции.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|Получает уникальный ID макета слайда.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|Получает имя макета слайда.|
||[shapes](/javascript/api/powerpoint/powerpoint.slidelayout#shapes)|Возвращает коллекцию фигур в макете слайда.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getCount__)|Получает количество макетов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItem_key_)|Получает макет с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemAt_index_)|Получает макет с использованием нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemOrNullObject_id_)|Получает макет с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|Получает уникальный ID мастера слайдов.|
||[макеты](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|Получает коллекцию макетов, предоставленных мастером слайдов для слайдов.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|Получает уникальное имя мастера слайдов.|
||[shapes](/javascript/api/powerpoint/powerpoint.slidemaster#shapes)|Возвращает коллекцию фигур в мастере слайдов.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getCount__)|Получает число мастеров слайдов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItem_key_)|Получает мастер слайдов с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemAt_index_)|Получает мастер слайдов с помощью нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemOrNullObject_id_)|Получает мастер слайдов с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|Получает уникальный ID тега.|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|Получает значение тега.|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_)|Добавляет новый тег в конце коллекции.|
||[delete (key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#delete_key_)|Удаляет тег с заданным `key` в этой коллекции.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getCount__)|Получает количество тегов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItem_key_)|Получает тег с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemAt_index_)|Получает тег с использованием нулевого индекса в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemOrNullObject_key_)|Получает тег с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[TextFrame](/javascript/api/powerpoint/powerpoint.textframe)|[autoSizeSetting](/javascript/api/powerpoint/powerpoint.textframe#autoSizeSetting)|Автоматические параметры размеров для текстового кадра.|
||[bottomMargin](/javascript/api/powerpoint/powerpoint.textframe#bottomMargin)|Представляет нижнее поле рамки с текстом (в пунктах).|
||[deleteText()](/javascript/api/powerpoint/powerpoint.textframe#deleteText__)|Удаляет весь текст в рамке с текстом.|
||[hasText](/javascript/api/powerpoint/powerpoint.textframe#hasText)|Указывает, содержит ли текстовый кадр текст.|
||[leftMargin](/javascript/api/powerpoint/powerpoint.textframe#leftMargin)|Представляет левое поле рамки с текстом (в пунктах).|
||[rightMargin](/javascript/api/powerpoint/powerpoint.textframe#rightMargin)|Представляет правое поле рамки с текстом (в пунктах).|
||[textRange](/javascript/api/powerpoint/powerpoint.textframe#textRange)|Представляет текст, присоединенный к фигуре в текстовой рамке, а также свойства и методы для операций с текстом.|
||[topMargin](/javascript/api/powerpoint/powerpoint.textframe#topMargin)|Представляет верхнее поле рамки с текстом (в пунктах).|
||[verticalAlignment](/javascript/api/powerpoint/powerpoint.textframe#verticalAlignment)|Представляет вертикальное выравнивание для рамки с текстом.|
||[wordWrap](/javascript/api/powerpoint/powerpoint.textframe#wordWrap)|Определяет, нарушаются ли строки автоматически, чтобы соответствовать тексту внутри фигуры.|
|[TextRange](/javascript/api/powerpoint/powerpoint.textrange)|[font](/javascript/api/powerpoint/powerpoint.textrange#font)|Возвращает `ShapeFont` объект, который представляет атрибуты шрифта для диапазона текста.|
||[getSubstring(start: number, length?: number)](/javascript/api/powerpoint/powerpoint.textrange#getSubstring_start__length_)|Возвращает объект `TextRange` для подстройки в заданный диапазон.|
||[paragraphFormat](/javascript/api/powerpoint/powerpoint.textrange#paragraphFormat)|Представляет формат абзаца в текстовом диапазоне.|
||[text](/javascript/api/powerpoint/powerpoint.textrange#text)|Представляет содержимое с обычным текстом в диапазоне текста.|
