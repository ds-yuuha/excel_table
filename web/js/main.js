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
        if (id == res) {
          $p.classList.add("options", "selected");
        } else {
          $p.classList.add("options");
        }
        const index = Object.keys(index_array).filter(function (k) {
          return index_array[k] == id[0];
        })[0];
        $p.textContent = index;

        fragment.appendChild($p);
        elements = indexList.appendChild(fragment);
      });
    });
  });

  eel.getSelectedItem()((result) => {
    let n = 0;
    let m = 0;
    result[0].forEach((r) => {
      id_name = "item_head" + n;
      if (n == 0) {
        if (document.getElementById("name")) {
          document.getElementById("name").value = r;
        }
      } else {
        if (document.getElementById(id_name)) {
          document.getElementById(id_name).value = r;
        }
      }
      if (itemList) {
        if (n > 0) {
          let $li = document.createElement("li");
          let $p1 = document.createElement("p");
          let $p2 = document.createElement("p");
          $p1.classList.add("item_tag");
          $p2.classList.add("item_content");
          $p1.textContent = "No." + n;
          $p2.textContent = r;
          $li.appendChild($p1);
          $li.appendChild($p2);
          $li.classList.add("items");

          item_fragment.appendChild($li);
          items = itemList.appendChild(item_fragment);
        }
      }
      n++;
    });
    result[1].forEach((r) => {
      setting_name = "setting" + m;
      if (m == 0) {
        if (document.getElementById("type")) {
          document.getElementById("type").value = r;
        }
      } else {
        if (document.getElementById(setting_name)) {
          document.getElementById(setting_name).value = r;
        }
      }
      m++;
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

if (document.getElementById("back")) {
  document.getElementById("back").addEventListener("click", () => {
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

targets = {};

setTimeout(function () {
  count = Object.keys(index_array).length;
  const find = setInterval(function () {
    for (let i = 0; i <= count; i++) {
      targets[i] = document.getElementsByClassName("options")[i];
      if (targets[i]) {
        clearInterval(find);
        targets[i].addEventListener("click", function (e) {
          const id = index_array[this.textContent];
          [...document.getElementsByClassName("options")].forEach((option) => {
            option.classList.remove("selected");
          });
          this.classList.add("selected");
          eel.changeIndex(id)((res) => {
            eel.getSelectedItem()((result) => {
              let n = 0;
              let m = 0;

              if (itemList) {
                itemList.replaceChildren();
              }
              item_fragment.replaceChildren();

              result[0].forEach((r) => {
                id_name = "item_head" + n;
                if (n == 0) {
                  if (document.getElementById("name")) {
                    document.getElementById("name").value = r;
                  }
                } else {
                  if (document.getElementById(id_name)) {
                    document.getElementById(id_name).value = r;
                  }
                }
                if (itemList) {
                  if (n > 0) {
                    let $li = document.createElement("li");
                    let $p1 = document.createElement("p");
                    let $p2 = document.createElement("p");
                    $p1.classList.add("item_tag");
                    $p2.classList.add("item_content");
                    $p1.textContent = "No." + n;
                    $p2.textContent = r;
                    $li.appendChild($p1);
                    $li.appendChild($p2);
                    $li.classList.add("items");

                    item_fragment.appendChild($li);
                    items = itemList.appendChild(item_fragment);
                  }
                }
                n++;
              });
              result[1].forEach((r) => {
                setting_name = "setting" + m;
                if (m == 0) {
                  if (document.getElementById("type")) {
                    document.getElementById("type").value = r;
                  }
                } else {
                  if (document.getElementById(setting_name)) {
                    document.getElementById(setting_name).value = r;
                  }
                }
                m++;
              });
            });
          });
        });
      }
    }
  }, 100);
}, 500);
