function callApi(endpoint, token) {
  var xhr = new XMLHttpRequest();

  // Setup our listener to process completed requests
  xhr.onreadystatechange = function () {

    // Only run if the request is complete
    if (xhr.readyState !== 4) return;

    // Process our return data
    if (xhr.status >= 200 && xhr.status < 300) {

      // What do when the request is successful
      console.log(JSON.parse(xhr.responseText));
      logMessage('Web API responded: Hello ' + JSON.parse(xhr.responseText)['name'] + '!');
    }
  };

  // Send the request using bearer token
  xhr.open('GET', endpoint);
  xhr.setRequestHeader('Authorization', 'Bearer ' + token);
  xhr.send();
}