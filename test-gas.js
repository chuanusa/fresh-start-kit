const API_URL = "https://script.google.com/macros/s/AKfycbwDuDK2BYwykf0Z-u2FNFwxqyu0NZbE4emYceSMAIa3oD5JRUB9zIzRHfbVxtHdEzfnlg/exec";

async function testFetch() {
    const formData = new URLSearchParams();
    formData.append("action", "getMonthlyDashboardData");
    formData.append("args", JSON.stringify([2026, 2]));

    console.log("Sending request to:", API_URL);
    try {
        const response = await fetch(API_URL, {
            method: "POST",
            body: formData,
            redirect: "follow",
        });

        console.log("Status:", response.status);
        console.log("Headers:", response.headers);

        const result = await response.text();
        console.log("Body length:", result.length);
        console.log("Body prefix:", result.substring(0, 500));
    } catch (error) {
        console.error("Fetch error:", error);
    }
}

testFetch();
