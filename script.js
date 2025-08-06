const selectedItems = [];

// Handle clicks on all products
document.querySelectorAll('.product').forEach(prod => {
  prod.addEventListener('click', () => {
    prod.classList.toggle('selected');
    
    const data = {
      rack: prod.dataset.rack,
      shelf: prod.dataset.shelf,
      name: prod.dataset.name
    };

    const index = selectedItems.findIndex(
      i => i.rack === data.rack && i.shelf === data.shelf && i.name === data.name
    );

    if (index >= 0) {
      // Already selected, remove it
      selectedItems.splice(index, 1);
    } else {
      // Not selected, add it
      selectedItems.push(data);
    }
  });
});

// Handle download button
document.getElementById('downloadExcel').addEventListener('click', () => {
  if (selectedItems.length === 0) {
    alert("Please select at least one item.");
    return;
  }

  // Prepare data rows
  const ws_data = [
    ["Rack", "Shelf", "Product"]
  ];

  selectedItems.forEach(item => {
    ws_data.push([item.rack, item.shelf, item.name]);
  });

  // Create worksheet and workbook
  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "SelectedItems");

  // Trigger download
  XLSX.writeFile(wb, "SelectedItems.xlsx");
});