const index_array = {};
const indexList = document.getElementById("index_list");

let itemList = "";
if (document.getElementById("item_list")) {
  itemList = document.getElementById("item_list");
}
const fragment = document.createDocumentFragment();
const item_fragment = document.createDocumentFragment();
let elements = "";
let items = "";

window.addEventListener("DOMContentLoaded", async () => {
  await eel.getIndexList()((res) => {
    const ids = res[0];
    const index_names = res[1];

    for (let i = 0; i < ids.length; i++) {
      index_array[index_names[i]] = ids[i];
    }

    ids.forEach((id) => {
      eel.getCurrentId()((res) => {
        let $p = document.createElement("li");
        $p.classList.add("options");
        const index = Object.keys(index_array).filter(function (k) {
          return index_array[k] == id[0];
        })[0];
        $p.textContent = index;

        fragment.appendChild($p);
        elements = indexList.appendChild(fragment);
      });
    });
  });
});

const link = (target) => {
  window.location.href = target;
};

const reload = () => {
  window.location.reload();
};

if (document.getElementById("register")) {
  document.getElementById("register").addEventListener("click", () => {
    link("register.html");
  });
}

if (document.getElementById("edit")) {
  document.getElementById("edit").addEventListener("click", () => {
    link("edit.html");
  });
}

if (document.getElementById("trim")) {
  document.getElementById("trim").addEventListener("click", () => {
    link("trim.html");
  });
}

if (document.getElementById("index")) {
  document.getElementById("index").addEventListener("click", () => {
    link("index.html");
  });
}

if (document.getElementById("manual")) {
  document.getElementById("manual").addEventListener("click", () => {
    link("manual.html");
  });
}

if (document.getElementById("reload")) {
  document.getElementById("reload").addEventListener("click", () => {
    reload();
  });
}
