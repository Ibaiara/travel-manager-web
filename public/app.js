fetch("/api/trips")
  .then(res => res.json())
  .then(trips => {
    const list = document.getElementById("trip-list");
    trips.forEach(t => {
      const li = document.createElement("li");
      li.textContent = `${t.destination} - ${t.status}`;
      list.appendChild(li);
    });
  });