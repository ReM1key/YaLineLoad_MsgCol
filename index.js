const lines = [
  '0 Центр Регионы',
  '1 Вип Центр Регионы',
  '1 Вип Центр Регионы Эксперимент',
  'Двойная оплата',
  'Пользователи Активный заказ',
  'ПользователиОтмены',
  '1 Отзывы Пользователи',
  '1 Финансы Пользователи',
  '1 Фрахт Пользователи',
  '1 ТАКСИ Почта Клиенты',
  '1 Почта Такси клиенты',
  'Пользователи Горящие',
  '1 Попрошайки',
  '1 Казахстан Пользователи',
  'Пользователи после проактива',
  'Такси СМС Сортировка',
  'Технический Проактив пользователи',
  'Финансы пользователи Тест',
  '2 Отзывы',
  '2 ФиныТех',
  '2 Почта',
  '2 Экстренные',
  '2 Неизвестные',
  'Негатив Пользователи',
  '2  Таски',
  '2 Перевод КГ',
  '2 Перевод КЗ КГ',
  '2 Пользователи Компенсации'
];

const standartList = [
  ['0 Центр + Регионы', lines[0]],
  ['1 Вип + Центр + Регионы', lines[1]],
  ['1 Вип + Центр + Регионы Эксперимент', lines[2]],
  ['Двойная оплата', lines[3]],
  ['Пользователи Активный заказ', lines[4]],
  ['Пользователи Отмены', lines[5]],
  ['1 · Отзывы · Пользователи', lines[6]],
  ['1 · Финансы · Пользователи', lines[7]],
  ['1 · Фрахт · Пользователи', lines[8]],
  ['1 · ТАКСИ Почта [Клиенты]', lines[9]],
  ['1 · Почта [Такси клиенты]', lines[10]],
  ['Пользователи Горящие', lines[11]],
  ['1 · Попрошайки', lines[12]],
  ['1 · Казахстан · Пользователи', lines[13]],
  ['Пользователи — после проактива', lines[14]],
  ['Такси: СМС Сортировка', lines[15]],
  ['[Технический] Проактив пользователи', lines[16]],
  ['Финансы пользователи Тест', lines[17]],
  ['2 Отзывы', lines[18]],
  ['2 Фины+Тех', lines[19]],
  ['2 · Почта', lines[20]],
  ['2 Экстренные', lines[21]],
  ['2 · Неизвестные списания', lines[22]],
  ['Негатив Пользователи', lines[23]],
  ['2 · Таски', lines[24]],
  ['2 · Перевод · КГ', lines[25]],
  ['2 · Перевод · КЗ · КГ', lines[26]],
  ['2 · Пользователи · Компенсации', lines[27]],
]

const button = document.getElementById('processButton');
const info_block = document.querySelector('.info-block');
const info_button = document.getElementById('info_button');

document.querySelector('.input-file input[type=file]').addEventListener('change', function() {
  const file = this.files[0];
  const fileNameElement = this.nextElementSibling;
  fileNameElement.textContent = file.name;
});

button.addEventListener('click', () => {
  const input = document.getElementById('loadJson');
  const file = input.files[0];

  const startDateInput = document.getElementById('startDataRule').value;
  const endDateInput = document.getElementById('endDataRule').value;
  const login = document.getElementById('loginTg').value;

  try {
    processJsonFile(file, (jsonData) => {
        takeBotMsgs(jsonData, startDateInput, endDateInput, login);
    });
    button.textContent = 'Генерирую эксель файл'
  } catch (error) {
      window.alert(`Вышла ошибочка:\n${error}`)
  }
});

const processJsonFile = (file, callback) => {
  const reader = new FileReader();

  reader.onload = (event) => {
      const data = event.target.result;
      const jsonData = JSON.parse(data);
      callback(jsonData);
  };

  reader.readAsText(file);
};

const takeBotMsgs = (data, startDate, endDate, login) => {
    const messages = data.messages; 
    let bot_msgs = messages.filter(el => el.from == login)
    
    switch (true) {
      case startDate !== "" && endDate !== "":
          bot_msgs = bot_msgs.filter(el => {
              return el.date.slice(0, 10) >= startDate && el.date.slice(0, 10) <= endDate;
          });
          break;
      case startDate !== "":
          bot_msgs = bot_msgs.filter(el => {
              return el.date.slice(0, 10) >= startDate;
          });
          break;
      case endDate !== "":
          bot_msgs = bot_msgs.filter(el => {
              return el.date.slice(0, 10) <= endDate;
          });
          break;
      default:
          break;
  }

    const filteredBot_msgs = bot_msgs.map(msg => {
        const {id, date, text} = msg;
        return { id, date, text };
    })

    splitLines(filteredBot_msgs)
}

const splitLines = (botMsgs) => {
  const result = botMsgs.map(el => {
    const lineQueue = el.text
    .replace(/\n - /g, ' - ')
    .split("\n")
    .slice(4)
    .map(el => el.split(" - "));

    const lq = lineQueue.map(line => {
      let fLine = line[0];

      standartList.forEach(l => {
        if (line[0] == l[0]) {
          fLine = l[1]
        }     
      })
      return [fLine, line[1]] 
    })

    const newObj = { ...el }

    lq.forEach(line => {
      newObj[line[0]] = line[1]
    })

    return newObj 
  })

  createExcelFile(result)
}

const createExcelFile = (data) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet 1');
  
    const columns = [
      { header: 'id', key: 'id', width: 10 },
      { header: 'date', key: 'date', width: 20 },
      { header: 'text', key: 'text', width: 50 },
    ]

    lines.forEach(el => {
      columns.push({
        header: el,
        key: el,
        width: 10
      });
    });

    worksheet.columns = columns;

    data.forEach((item) => {
      worksheet.addRow(item);
    });
  
    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      document.body.appendChild(a);
      a.href = url;
      a.download = 'output.xlsx';
      a.click();
      window.URL.revokeObjectURL(url);
      button.textContent = 'Готово'
      info_block.style.display = 'grid';
  });
};


info_button.addEventListener('click', () => {
  const container_form = document.querySelector('.container-form');
  const container_info = document.querySelector('.container-info');

  container_form.style.display = 'none';
  container_info.style.display = 'grid';
})