form.addEventListener("submit", async e => {
  e.preventDefault();
  const res = await fetch("/api/process", {
    method: "POST",
    body: new FormData(form)
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const json = await res.json();
  console.log(json);
});