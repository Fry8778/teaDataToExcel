function silpoExportToExcel() {     
    const productCards = document.querySelectorAll('.product-card__body');
    const filteredProducts = Array.from(productCards).filter(productCard => {
        const productNameElement = productCard.querySelector('.product-card__title');           
        const productName = productNameElement ? productNameElement.innerText.toLowerCase() : ''; // Проверка существования элемента
            return  productName.includes("чай") ||
                    productName.includes("суміш") ||
                    productName.includes("чай чорний") ||        
                    productName.includes("чай трав'яний")  ||
                    productName.includes("чай фруктовий") ||
                    productName.includes("чай фруктово-трав'яний") ||
                    productName.includes("напій фруктово-трав'яний") ||
                    productName.includes("суміш фруктово-ягідна") ||
                    productName.includes("чай фруктово-ягідний") ||
                    productName.includes("суміш трав") ||
                    productName.includes("чай зелений")  ||
                    productName.includes("чайні набори") ||
                    productName.includes("чайний напій") ||                    
                    productName.includes("чай чорний і зелений") ||
                    productName.includes("чай бірюзовий") ||
                    productName.includes("чай гречаний") ||
                    productName.includes("фіточай") ||
                    productName.includes("напій на основі екстракту чорного чаю") ||
                    productName.includes("набір чаю") ||
                    productName.includes("набір чаїв"); 
    });

    const data = [[ 'Название товара',            
                    'Цена товара(текущая цена)', 
                    'Вес товара',     
                    'Цена товара без скидки(старая цена)',
                    'Процент скидки(%)']];

    filteredProducts.forEach((productCard) => {
        const productNameElements = productCard.querySelectorAll('.product-card__title'); 
        const priceElement = productCard.querySelector('.ft-whitespace-nowrap.ft-text-22.ft-font-bold');
        const weightElement = productCard.querySelector('.ft-typo-14-semibold.xl\\:ft-typo-16-semibold > span');    

        // Проверяем наличие скидки
        const specialPriceElement = productCard.querySelector('.ft-line-through.ft-text-black-87.ft-typo-14-regular.xl\\:ft-typo');       
        const discountPercentageElement = productCard.querySelector('.product-card-price__sale');

        console.log("productNameElements:", productNameElements);        
        console.log("priceElement:", priceElement);
        console.log("weightElement:", weightElement);
        
        console.log("specialPriceElement:", specialPriceElement);        
        console.log("discountPercentageElement:", discountPercentageElement); 

        if (!specialPriceElement || !discountPercentageElement) {
            // Если элементов .ft-line-through.ft-text-black-87.ft-typo-14-regular.xl\\:ft-typo и .product-card-price__sale нет внутри .product-card-price__old,
            // значит, товар не имеет скидки
            const price = priceElement ? priceElement.innerText.trim() || '' : '';  
            const weight = weightElement ? weightElement.innerText.trim() || '' : '';  
            const productName = Array
                .from(productNameElements)
                .map(element => element.innerText.trim() || '')
                .join(' ');     
            const specialPrice = '';  // Пустое значение для товаров без скидки
            const discountPercentage = '';     // Пустое значение для товаров без скидки
            data.push([ productName,    
                        price,
                        weight,
                        specialPrice,                     
                        discountPercentage]);
        } else {
            // Если элементы .ft-line-through.ft-text-black-87.ft-typo-14-regular.xl\\:ft-typo и .product-card-price__sale найдены, значит, товар имеет скидку
            const productName = Array
                .from(productNameElements)
                .map(element => element.innerText.trim() || '')
                .join(' ');     
            const price = priceElement ? priceElement.innerText.trim() || '' : '';  
            const weight = weightElement ? weightElement.innerText.trim() || '' : '';  
            const specialPrice = specialPriceElement ? specialPriceElement.innerText.trim() || '' : '';               
            const discountPercentage = discountPercentageElement ? discountPercentageElement.innerText.trim() || '' : '';             
            data.push([ productName,    
                        price,
                        weight,
                        specialPrice,                     
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

export { silpoExportToExcel };