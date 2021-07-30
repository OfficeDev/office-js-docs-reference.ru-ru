| Класс | Поля | Описание |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[форматирование](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Указывает форматирование, которое необходимо использовать во время вставки слайда.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceSlideIds)|Указывает слайды из исходных презентаций, которые будут вставлены в текущую презентацию.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetSlideId)|Указывает, где в презентации будут вставлены новые слайды.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint. InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertSlidesFromBase64_base64File__options_)|Вставляет указанные слайды из презентации в текущую презентацию.|
||[слайды](/javascript/api/powerpoint/powerpoint.presentation#slides)|Возвращает упорядоченную коллекцию слайдов в презентации.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete__)|Удаляет слайд из презентации.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Получает уникальный ID слайда.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getCount__)|Получает количество слайдов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getItem_key_)|Получает слайд с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_)|Получает слайд с использованием нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidecollection#getItemOrNullObject_id_)|Получает слайд с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
