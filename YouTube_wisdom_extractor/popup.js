document.getElementById("extractButton").addEventListener("click", () => {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      const activeTab = tabs[0];
      const youtubeUrl = activeTab.url;
  
      fetch(`http://localhost:8501/extract`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({ url: youtubeUrl })
      })
      .then(response => response.json())
      .then(data => {
        console.log(data.message);
      })
      .catch(error => {
        console.error('Error:', error);
      });
    });
  });
  