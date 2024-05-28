function ashanExportToExcel() {     
    const productCards = document.querySelectorAll('.item_container__17WaH');
    const filteredProducts = Array.from(productCards).filter(productCard => {
        const productNameElement = productCard.querySelector('.item_data__name__3bRJz');           
        const productName = productNameElement ? productNameElement.innerText.toLowerCase() : ''; // Проверка существования элемента
            return  productName.includes("чай") ||
                    productName.includes("суміш") ||
                    productName.includes("суміш чаїв") ||
                    productName.includes("суміш чаю") ||
                    productName.includes("колекція чаю") ||
                    productName.includes("чай чорний") ||        
                    productName.includes("чорний чай") ||        
                    productName.includes("чорний і зелений чай") ||        
                    productName.includes("бленд чорного та зеленого чаю") ||        
                    productName.includes("чай трав'яний")  ||
                    productName.includes("трав'яний чай")  ||
                    productName.includes("чай фруктовий") ||
                    productName.includes("чай фруктово-трав'яний") ||
                    productName.includes("напій фруктово-трав'яний") ||
                    productName.includes("суміш фруктово-ягідна") ||
                    productName.includes("чай фруктово-ягідний") ||
                    productName.includes("чай фруктово-медовий") ||
                    productName.includes("чай квітковий та ягідний") ||
                    productName.includes("чай плодово-ягідний та квітковий") ||
                    productName.includes("суміш трав") ||
                    productName.includes("чай зелений")  ||
                    productName.includes("чайні набори") ||
                    productName.includes("чайний напій") ||                    
                    productName.includes("чай чорний і зелений") ||
                    productName.includes("чай бірюзовий") ||
                    productName.includes("чай гречаний") ||
                    productName.includes("фіточай") ||
                    productName.includes("фільтр–пакети для чаю") ||
                    productName.includes("напій на основі екстракту чорного чаю") ||
                    productName.includes("напій на основі зеленого чаю") ||
                    productName.includes("подарунковий набір чаю") ||
                    productName.includes("набір-асорті чаїв") ||
                    productName.includes("набір-асорті чаю") ||
                    productName.includes("набір чаю") ||
                    productName.includes("набір чорного чаю") ||
                    productName.includes("набір чаїв");  
    });

    const data = [[ 'Название товара',            
                    'Цена товара(текущая цена)',                     
                    'Цена товара с учетом скидки(текущая цена)',
                    'Старая цена товара(цена без скидки)',
                    'Процент скидки(%)']];

    filteredProducts.forEach((productCard) => {
        const productNameElements = productCard.querySelectorAll('.item_data__name__3bRJz'); 
        const priceElement = productCard.querySelector('.item_price__sEYUp > span.item_price_value_actual__2hWfO');       

        // Проверяем наличие скидки
        const specialPriceElement = productCard.querySelector('.item_price__sEYUp > span.item_price_value_actual__2hWfO');  
        const salePriceElement = productCard.querySelector('.item_price__sEYUp > span.item_price_value_old__H2oLx');     
        const discountPercentageElement = productCard.querySelector('.item_percent__1EHD0');

        console.log("productNameElements:", productNameElements);        
        console.log("priceElement:", priceElement);     
        
        console.log("specialPriceElement:", specialPriceElement);      
        console.log("salePriceElement:", salePriceElement);   
        console.log("discountPercentageElement:", discountPercentageElement); 

        if (!specialPriceElement || !salePriceElement || !discountPercentageElement) {
            // Если элементов .ft-line-through.ft-text-black-87.ft-typo-14-regular.xl\\:ft-typo и .product-card-price__sale нет внутри .product-card-price__old,
            // значит, товар не имеет скидки
            const price = priceElement ? priceElement.innerText.trim() || '' : '';          
            const productName = Array
                .from(productNameElements)
                .map(element => element.innerText.trim() || '')
                .join(' ');     
            const specialPrice = '';  // Пустое значение для товаров без скидки
            const salePrice = '';     // Пустое значение для товаров без скидки
            const discountPercentage = '';     // Пустое значение для товаров без скидки
            data.push([ productName,    
                        price,
                        specialPrice,
                        salePrice,                     
                        discountPercentage]);
        } else {
            // Если элементы .ft-line-through.ft-text-black-87.ft-typo-14-regular.xl\\:ft-typo и .product-card-price__sale найдены, значит, товар имеет скидку
            const productName = Array
                .from(productNameElements)
                .map(element => element.innerText.trim() || '')
                .join(' ');     
            const price = '';             
            const specialPrice = specialPriceElement ? specialPriceElement.innerText.trim() || '' : ''; 
            const salePrice = salePriceElement ? salePriceElement.innerText.trim() || '' : '';              
            const discountPercentage = discountPercentageElement ? discountPercentageElement.innerText.trim() || '' : '';             
            data.push([ productName,    
                        price,
                        specialPrice,
                        salePrice,                     
                        discountPercentage]);
        }
    });

    if (data.length <= 1) {
        alert("На странице нет данных для экспорта в Excel.");
        return;
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, "data.xlsx");
}

export { ashanExportToExcel };