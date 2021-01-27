| Класс | Поля | Описание |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[formatting](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Указывает форматирование, которое будет применяться во время вставки слайда.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceslideids)|Указывает слайды из презентации источника, которые будут вставлены в текущую презентацию.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetslideid)|Указывает место вставки новых слайдов в презентацию.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|Вставляет указанные слайды из презентации в текущую презентацию.|
||[slides](/javascript/api/powerpoint/powerpoint.presentation#slides)|Возвращает упорядоченную коллекцию слайдов в презентации.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete--)|Удаляет слайд из презентации.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Получает уникальный ИД слайда.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getcount--)|Получает количество слайдов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitem-key-)|Получает слайд по уникальному ИД.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemat-index-)|Получает слайд с помощью индекса на основе нуля в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemornullobject-id-)|Получает слайд с использованием уникального ИД.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
