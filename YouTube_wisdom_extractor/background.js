chrome.runtime.onInstalled.addListener(() => {
    chrome.contextMenus.create({
      id: "extractWisdom",
      title: "Extract Wisdom",
      contexts: ["page"]
    });
  });
  
  chrome.contextMenus.onClicked.addListener((info, tab) => {
    if (info.menuItemId === "extractWisdom") {
      const youtubeUrl = info.pageUrl;
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
    }
  });
  