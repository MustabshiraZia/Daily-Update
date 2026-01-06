const people = ["Alex", "Jamie", "Taylor", "Morgan"];

const channel = new BroadcastChannel("daily-update");

let state = people.map(name => ({
  name,
  mood: "",
  success: "",
  discovery: "",
  concern: ""
}));

const board = document.getElementById("board");

function render() {
  board.innerHTML = "";
  state.forEach((row, i) => {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td><strong>${row.name}</strong></td>
      <td>
        <select data-i="${i}" data-f="mood">
          <option value=""></option>
          <option>😞</option>
          <option>😐</option>
          <option>🙂</option>
          <option>😄</option>
          <option>🚀</option>
        </select>
      </td>
      <td><input data-i="${i}" data-f="success" value="${row.success}" /></td>
      <td><input data-i="${i}" data-f="discovery" value="${row.discovery}" /></td>
      <td><input data-i="${i}" data-f="concern" value="${row.concern}" /></td>
    `;

    tr.querySelector("select").value = row.mood;
    board.appendChild(tr);
  });
}

board.addEventListener("input", e => {
  const i = e.target.dataset.i;
  const f = e.target.dataset.f;
  if (!f) return;

  state[i][f] = e.target.value;
  channel.postMessage(state);
});

channel.onmessage = e => {
  state = e.data;
  render();
};

render();
