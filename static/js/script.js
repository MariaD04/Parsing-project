document.getElementById('parseBtn').addEventListener('click', async function() {
    const btn = this;
    const resultDiv = document.getElementById('result');
    const loadingDiv = document.getElementById('loading');
    
    // Получаем значения из полей ввода
    const docxIn = document.getElementById('docx_in').value.trim();
    const xlsxIn = document.getElementById('xlsx_in').value.trim();
    const docxOut = document.getElementById('docx_out').value.trim();
    const xlsxOut = document.getElementById('xlsx_out').value.trim();
    
    // Проверяем, что все поля заполнены
    if (!docxIn || !xlsxIn || !docxOut || !xlsxOut) {
        resultDiv.innerHTML = '<p class="error">Пожалуйста, заполните все поля!</p>';
        return;
    }
    
    btn.disabled = true;
    loadingDiv.style.display = 'block';
    resultDiv.innerHTML = '';
    
    try {
        const apiUrl = 'http://localhost:8000/parse';
        
        const params = new URLSearchParams({
            docx_in: docxIn,
            xlsx_in: xlsxIn,
            docx_out: docxOut,
            xlsx_out: xlsxOut
        });
        
        const response = await fetch(`${apiUrl}?${params.toString()}`);
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.detail || `Ошибка HTTP: ${response.status}`);
        }
        
        const data = await response.json();
        
        resultDiv.innerHTML = `
            <h3>Результат обработки:</h3>
            <p><strong>Статус:</strong> ${data.status}</p>
            <p><strong>DOCX файл сохранен:</strong> ${data.output_file}</p>
            <p><strong>XLSX файл сохранен:</strong> ${data.xlsx_output}</p>
            <h4>Краткое содержание DOCX:</h4>
            <p>${data.result.docx_summary}</p>
            <h4>Краткое содержание XLSX:</h4>
            <p>${data.result.xlsx_summary}</p>
        `;
    } catch (error) {
        resultDiv.innerHTML = `
            <p class="error">Произошла ошибка:</p>
            <p>${error.message}</p>
        `;
    } finally {
        btn.disabled = false;
        loadingDiv.style.display = 'none';
    }
});