const fs = require('fs');

async function testGas() {
    const API_URL = "https://script.google.com/macros/s/AKfycbwDuDK2BYwykf0Z-u2FNFwxqyu0NZbE4emYceSMAIa3oD5JRUB9zIzRHfbVxtHdEzfnlg/exec";
    const formData = new URLSearchParams();
    formData.append("action", "getMonthlyDashboardData");
    formData.append("args", JSON.stringify([2026, 2]));

    try {
        const response = await fetch(API_URL, {
            method: "POST",
            body: formData,
        });
        const text = await response.text();
        fs.writeFileSync('gas_output.json', text, 'utf8');
        console.log("Status:", response.status);
        console.log("Length:", text.length);
        console.log("First 100 chars:", text.substring(0, 100));
    } catch (err) {
        console.error(err);
    }
}

testGas();
