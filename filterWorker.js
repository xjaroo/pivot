self.onmessage = function (e) {
  const { values, filter } = e.data;
  // Get unique values
  const set = new Set(values);
  let unique = Array.from(set);
  // Filter
  if (filter && filter.trim() !== "") {
    const f = filter.trim().toLowerCase();
    unique = unique.filter((v) => String(v).toLowerCase().includes(f));
  }
  // Sort
  unique.sort((a, b) => String(a).localeCompare(String(b)));
  self.postMessage({ unique });
};
