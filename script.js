function analyzeExcelFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
  
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const cellValueB2 = worksheet.B2 ? worksheet.B2.v : 0;
        const cellValueC2 = worksheet.C2 ? worksheet.C2.v : 0;
        const cellValueD2 = worksheet.D2 ? worksheet.D2.v : 0;
  
        let resultText = '';
        if (cellValueB2 >= 0) {
          resultText += `Количество убийств выросло на ${cellValueB2}%\n\n`;
        } else {
          resultText += `Количество убийств снизилось на ${Math.abs(cellValueB2)}%\n\n`;
        }
  
        if (cellValueC2 >= 0) {
          resultText += `Количество грабежей выросло на ${cellValueC2}%\n\n`;
        } else {
          resultText += `Количество грабежей снизилось на ${Math.abs(cellValueC2)}%\n\n`;
        }
  
        if (cellValueD2 >= 0) {
          resultText += `Количество ДТП выросло на ${cellValueD2}%\n\n`;
        } else {
          resultText += `Количество ДТП снизилось на ${Math.abs(cellValueD2)}%\n\n`;
        }
  
        resolve(resultText);
      };
  
      reader.onerror = (event) => {
        reject(event.target.error);
      };
  
      reader.readAsArrayBuffer(file);
    });
  }
  
  function createWordDocument(resultText) {
    const link = document.createElement('a');
    const content = 'Результат анализа:\n\n' + resultText;
    const encodedContent = encodeURIComponent(content);
    link.href = 'data:text/plain;charset=utf-8,' + encodedContent;
    link.download = 'результат.txt';
    link.textContent = 'Скачать результат';
    return link;
  }
  