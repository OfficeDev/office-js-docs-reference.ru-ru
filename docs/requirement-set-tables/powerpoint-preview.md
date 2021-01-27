| Класс | Поля | Описание |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutid)|Указывает ИД макета слайда, который будет использоваться для нового слайда.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slidemasterid)|Указывает ИД хозяина слайда, который будет использоваться для нового слайда.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slidemasters)|Возвращает коллекцию `SlideMaster` объектов, которые находятся в презентации.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[id](/javascript/api/powerpoint/powerpoint.shape#id)|Получает уникальный ИД фигуры.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getcount--)|Получает количество фигур в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitem-key-)|Получает фигуру по уникальному ИД.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemat-index-)|Получает фигуру с помощью индекса на основе нуля в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemornullobject-id-)|Получает фигуру по уникальному ИД.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[layout](/javascript/api/powerpoint/powerpoint.slide#layout)|Получает макет слайда.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Возвращает коллекцию фигур на слайде.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slidemaster)|Получает `SlideMaster` объект, который представляет содержимое слайда по умолчанию.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint.AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add-options-)|Добавляет новый слайд в конец коллекции.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|Получает уникальный ИД макета слайда.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|Получает имя макета слайда.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getcount--)|Получает количество макетов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitem-key-)|Получает макет по уникальному ИД.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemat-index-)|Получает макет с помощью индекса на основе нуля в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemornullobject-id-)|Получает макет по уникальному ИД.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|Получает уникальный ИД хозяина слайда.|
||[layouts](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|Получает коллекцию макетов, предоставленных мастером слайдов для слайдов.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|Получает уникальное имя мастера слайдов.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getcount--)|Получает количество мастеров слайдов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitem-key-)|Получает хозяин слайда по уникальному ИД.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemat-index-)|Получает хозяин слайда с помощью индекса на основе нуля в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemornullobject-id-)|Получает мастер слайдов по уникальному ИД.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
