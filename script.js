document.getElementById('submitBtn').addEventListener('click', saveToExcel);

async function saveToExcel() {
    const requiredFields = [
        { id: 'auto', name: 'Item Code (Auto Number)' },
        { id: 'name', name: 'Item Name' },
        { id: 'descrpt', name: 'Item Description' },
        { id: 'remark', name: 'Item Remarks' },
        { id: 'remarks', name: 'Remarks' }
    ];

    let isValid = true;
    let errorMessage = "Please fill the following required fields:\n";

    // Validate required fields
    requiredFields.forEach(field => {
        const element = document.getElementById(field.id);
        if (!element.value.trim()) {
            isValid = false;
            errorMessage += `${field.name}\n`;
        }
    });

    if (!isValid) {
        alert(errorMessage);
        return;
    }

    // Collect form data
    const formData = {
        auto: document.getElementById('auto').value || 'N/A',
        manual: document.getElementById('manual').value || 'N/A',
        itemImage: document.getElementById('itemimage').value || 'N/A',
        itemLabel: document.getElementById('itemlabel').value || 'N/A',
        drawing: document.getElementById('drawing').value || 'N/A',
        datasheet: document.getElementById('datasheet').value || 'N/A',
        datasheet1: document.getElementById('datasheet1').value || 'N/A',
        status: document.getElementById('status').value || 'N/A',
        name: document.getElementById('name').value || 'N/A',
        type: document.getElementById('type').value || 'N/A',
        productid: document.getElementById('productid').value || 'N/A',
        bom: document.getElementById('bom').value || 'N/A',
        cat: document.getElementById('cat').value || 'N/A',
        subcat: document.getElementById('subcat').value || 'N/A',
        subcat1: document.getElementById('subcat1').value || 'N/A',
        group: document.getElementById('group').value || 'N/A',
        source: document.getElementById('source').value || 'N/A',
        stores: document.getElementById('stores').value || 'N/A',
        storeloc: document.getElementById('storeloc').value || 'N/A',
        floor: document.getElementById('floor').value || 'N/A',
        descrpt: document.getElementById('descrpt').value || 'N/A',
        colour: document.getElementById('colour').value || 'N/A',
        weight: document.getElementById('weight').value || 'N/A',
        dimension: document.getElementById('dimension').value || 'N/A',
        wattage: document.getElementById('wattage').value || 'N/A',
        wattage1: document.getElementById('wattage1').value || 'N/A',
        wattage2: document.getElementById('wattage2').value || 'N/A',
        voltage: document.getElementById('voltage').value || 'N/A',
        voltage1: document.getElementById('voltage1').value || 'N/A',
        voltage2: document.getElementById('voltage2').value || 'N/A',
        current: document.getElementById('current').value || 'N/A',
        current1: document.getElementById('current1').value || 'N/A',
        current2: document.getElementById('current2').value || 'N/A',
        ip: document.getElementById('ip').value || 'N/A',
        cri: document.getElementById('cri').value || 'N/A',
        beam: document.getElementById('beam').value || 'N/A',
        alt: document.getElementById('alt').value || 'N/A',
        alt1: document.getElementById('alt1').value || 'N/A',
        alt2: document.getElementById('alt2').value || 'N/A',
        quantity: document.getElementById('quantity').value || 'N/A',
        unit1: document.getElementById('unit1').value || 'N/A',
        make: document.getElementById('make').value || 'N/A',
        ref: document.getElementById('ref').value || 'N/A',
        order: document.getElementById('order').value || 'N/A',
        proname: document.getElementById('proname').value || 'N/A',
        proid: document.getElementById('proid').value || 'N/A',
        drawingno: document.getElementById('drawingno').value || 'N/A',
        drawingset: document.getElementById('drawingset').value || 'N/A',
        cate: document.getElementById('cate').value || 'N/A',
        itemmaterial: document.getElementById('itemmaterial').value || 'N/A',
        itemtype: document.getElementById('itemtype').value || 'N/A',
        price: document.getElementById('price').value || 'N/A',
        big: document.getElementById('big').value || 'N/A',
        dot: document.getElementById('dot').value || 'N/A',
        cc: document.getElementById('cc').value || 'N/A',
        remark: document.getElementById('remark').value || 'N/A',
        remarks: document.getElementById('remarks').value || 'N/A'
    };

    // Prepare data for Excel
    const sheetData = [
        Object.keys(formData),
        Object.values(formData)
    ];

    // Create Excel workbook and sheet
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'FormData');

    // Convert workbook to Blob
    const excelBlob = XLSX.write(workbook, { bookType: 'xlsx', type: 'blob' });
}
