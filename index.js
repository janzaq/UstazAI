import { Telegraf, Markup } from 'telegraf';
import { GoogleGenAI } from '@google/genai';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType } from 'docx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import http from 'http';

// Токен бота и ключи Gemini из переменных окружения
const BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
if (!BOT_TOKEN) {
  console.error('ОШИБКА: TELEGRAM_BOT_TOKEN не задан!');
}

// 1. МЕНЕДЖЕР КЛЮЧЕЙ С АВТОПЕРЕКЛЮЧЕНИЕМ (FAILOVER)
class KeyPool {
  constructor() {
    const rawKeys = process.env.GEMINI_API_KEYS || process.env.GEMINI_API_KEY || '';
    const keysArray = rawKeys.split(',').map((k) => k.trim()).filter(Boolean);

    this.keys = keysArray.map((key, i) => ({
      id: `key-${i + 1}`,
      name: `Gemini-Key-${i + 1}`,
      key: key,
      status: 'ACTIVE',
      rpmUsed: 0,
      rpmLimit: 15,
      rpdUsed: 0,
      rpdLimit: 1500,
      cooldownUntil: null,
    }));

    // Сброс счетчика RPM раз в минуту
    setInterval(() => {
      const now = Date.now();
      for (const k of this.keys) {
        k.rpmUsed = 0;
        if (k.cooldownUntil && now >= k.cooldownUntil) {
          k.status = 'ACTIVE';
          k.cooldownUntil = null;
        }
      }
    }, 60000);
  }

  getActive() {
    const now = Date.now();
    for (const k of this.keys) {
      if (k.cooldownUntil && now >= k.cooldownUntil) {
        k.status = 'ACTIVE';
        k.cooldownUntil = null;
      }
    }
    const available = this.keys.find((k) => k.status === 'ACTIVE' && k.rpmUsed < k.rpmLimit);
    return available || this.keys[0];
  }

  handleFailover(keyId) {
    const failed = this.keys.find((k) => k.id === keyId);
    if (failed) {
      failed.status = 'COOLING_DOWN';
      failed.cooldownUntil = Date.now() + 60000; // 60 секунд отдыха
    }
    return this.getActive();
  }

  getQuotaText() {
    const active = this.getActive();
    if (!active) return '📊 Лимит: Ключи не настроены';
    const remainingRpm = Math.max(0, active.rpmLimit - active.rpmUsed);
    const activeCount = this.keys.filter((k) => k.status === 'ACTIVE').length;
    return `📊 **Бот лимиті / Остаток лимита бота:**\n• Белсенді кілт: \`${active.name}\`\n• Қалған минуттық лимит: **${remainingRpm} / ${active.rpmLimit} RPM**\n• Күндік қалдық: **${active.rpdLimit - active.rpdUsed} / ${active.rpdLimit} RPD**\n• Жұмыс істеп тұрған кілттер: **${activeCount} / ${this.keys.length}**`;
  }
}

const keyPool = new KeyPool();
const bot = new Telegraf(BOT_TOKEN || 'DUMMY_TOKEN');

// Память сессий пользователей
const userSessions = new Map();

function getSession(chatId) {
  if (!userSessions.has(chatId)) {
    userSessions.set(chatId, {
      step: 'SELECT_LANG',
      lang: 'kk',
      fullName: '',
      collegeName: '',
      history: [],
    });
  }
  return userSessions.get(chatId);
}

// 2. ОНБОРДИНГ: СТАРТ И ВЫБОР ЯЗЫКА
bot.start(async (ctx) => {
  const session = getSession(ctx.chat.id);
  session.step = 'SELECT_LANG';

  await ctx.reply(
    `👋 **Сәлеметсіз бе, құрметті колледж оқытушысы! / Здравствуйте, уважаемый преподаватель колледжа!**\n\n` +
      `Мен — ҚР **ГОСО ТиПО (МЖМБС ТжКБ)** талаптары бойынша поурочные планы (КСП), дескрипторлар мен раздаточный материалдарды дайындайтын интеллектуалды ботпын.\n\n` +
      `Жұмыс тілін таңдаңыз / Выберите язык документов:`,
    Markup.inlineKeyboard([
      [Markup.button.callback('🇰🇿 Қазақ тілі', 'lang_kk'), Markup.button.callback('🇷🇺 Русский язык', 'lang_ru')],
    ])
  );
});

// Кнопки выбора языка
bot.action('lang_kk', async (ctx) => {
  const session = getSession(ctx.chat.id);
  session.lang = 'kk';
  session.step = 'ENTER_NAME';
  await ctx.answerCbQuery();
  await ctx.reply('🇰🇿 **Қазақ тілі таңдалды.**\n\nҚұжаттарға ресми жазылуы үшін **толық Аты-жөніңізді (ФИО)** жазыңыз:\n*(Мысалы: Ахметов Бауыржан Серікұлы)*');
});

bot.action('lang_ru', async (ctx) => {
  const session = getSession(ctx.chat.id);
  session.lang = 'ru';
  session.step = 'ENTER_NAME';
  await ctx.answerCbQuery();
  await ctx.reply('🇷🇺 **Выбран русский язык.**\n\nДля официального оформления документов укажите ваше **полное ФИО**:\n*(Например: Сериков Данияр Болатович)*');
});

// 3. ОБРАБОТКА ТЕКСТА: ШАГИ ФИО, КОЛЛЕДЖ И ГЕНЕРАЦИЯ
bot.on('text', async (ctx) => {
  const session = getSession(ctx.chat.id);
  const text = ctx.message.text.trim();

  // Шаг ввода ФИО
  if (session.step === 'ENTER_NAME') {
    session.fullName = text;
    session.step = 'ENTER_COLLEGE';
    const msg =
      session.lang === 'kk'
        ? `Рахмет, құрметті **${session.fullName}**!\n\nЕнді қай **колледжде** сабақ беретініңізді жазыңыз:\n*(Мысалы: Алматы мемлекеттік политехникалық колледжі)*`
        : `Спасибо, уважаемый(ая) **${session.fullName}**!\n\nУкажите название вашего **колледжа**:\n*(Например: Высший политехнический колледж)*`;
    return ctx.reply(msg);
  }

  // Шаг ввода Колледжа
  if (session.step === 'ENTER_COLLEGE') {
    session.collegeName = text;
    session.step = 'COMPLETED';

    const welcome =
      session.lang === 'kk'
        ? `🎓 **Профиль сәтті сақталды!**\n👤 Оқытушы: ${session.fullName}\n🏛 Колледж: ${session.collegeName}\n\nЕнді сабақ тақырыбын жазыңыз немесе төмендегі батырмаларды қолданыңыз:`
        : `🎓 **Профиль успешно сохранен!**\n👤 Преподаватель: ${session.fullName}\n🏛 Колледж: ${session.collegeName}\n\nТеперь введите тему занятия или используйте готовые кнопки:`;

    await ctx.reply(
      welcome,
      Markup.keyboard([
        ['📚 Сабақ жоспарын жасау', '📑 КТП / РУП Колледжа'],
        ['🔑 Лимиттер мен Ключи', '🌐 Тілді ауыстыру / Сменить язык'],
      ]).resize()
    );
    return;
  }

  if (text === '🌐 Тілді ауыстыру / Сменить язык') {
    session.step = 'SELECT_LANG';
    return ctx.reply(
      'Тілді таңдаңыз / Выберите язык:',
      Markup.inlineKeyboard([
        [Markup.button.callback('🇰🇿 Қазақ тілі', 'lang_kk'), Markup.button.callback('🇷🇺 Русский язык', 'lang_ru')],
      ])
    );
  }

  if (text === '🔑 Лимиттер мен Ключи' || text === '/keys') {
    return ctx.reply(keyPool.getQuotaText());
  }

  // Генерация сабақ жоспары (КСП + Дескрипторы + Раздатка)
  await ctx.sendChatAction('typing');
  const waitMsg = await ctx.reply('⏳ ҚР ГОСО ТиПО стандарты бойынша құжаттар топтамасы әзірленуде...');

  const prompt = `Сен — ҚР ГОСО ТиПО (Техникалық және кәсіптік білім беру) бойынша сарапшысың.
Мұғалім: ${session.fullName || 'Оқытушы'}
Колледж: ${session.collegeName || 'Колледж'}
Тіл: ${session.lang}
Тақырып: ${text}

Тапсырма: Сабақ жоспарын құрастырып, тек қана келесідей таза JSON форматында қайтар:
{
  "subject": "Модуль/Пән атауы",
  "grade": "Курс және топ",
  "topic": "${text}",
  "objective": "Оқыту нәтижесі (ОН)",
  "table": [
    {"stage": "Ұйымдастыру", "time": "5 мин", "teacher": "Сәлемдесу, түгендеу", "student": "Дайындық", "assessment": "Ауызша мадақтау"},
    {"stage": "Жаңа білім", "time": "25 мин", "teacher": "Теорияны түсіндіру", "student": "Конспект жазу", "assessment": "Сұрақ-жауап"},
    {"stage": "Тәжірибелік кезең", "time": "40 мин", "teacher": "А, В, С тапсырмаларын беру", "student": "Тапсырманы орындау", "assessment": "10 балдық дескриптор"},
    {"stage": "Қорытынды", "time": "10 мин", "teacher": "Рефлексия мен баға қою", "student": "Өзін-өзі бағалау", "assessment": "Кері байланыс"}
  ]
}`;

  let parsedData = null;
  let attempts = 0;

  // Цикл с автоматическим переключением ключей (Failover)
  while (attempts < 3) {
    const activeKey = keyPool.getActive();
    if (!activeKey || !activeKey.key) break;
    activeKey.rpmUsed++;
    activeKey.rpdUsed++;

    try {
      const ai = new GoogleGenAI({ apiKey: activeKey.key });
      const res = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: [{ role: 'user', parts: [{ text: prompt }] }],
        config: { responseMimeType: 'application/json' },
      });
      parsedData = JSON.parse(res.text || '{}');
      break;
    } catch (err) {
      console.log(`Ключ ${activeKey.name} дал сбой, переключаем...`, err.message);
      keyPool.handleFailover(activeKey.id);
      attempts++;
    }
  }

  if (!parsedData || !parsedData.table) {
    return ctx.reply('⚠️ Қате орын алды немесе барлық кілттердің лимиті бітті. Біраздан соң қайталаңыз.');
  }

  // Генерация DOCX
  const rows = [
    new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Кезеңі', bold: true })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Уақыт', bold: true })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Оқытушы әрекеті', bold: true })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Студент әрекеті', bold: true })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Бағалау', bold: true })] })] }),
      ],
    }),
    ...parsedData.table.map(
      (r) =>
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph(r.stage)] }),
            new TableCell({ children: [new Paragraph(r.time)] }),
            new TableCell({ children: [new Paragraph(r.teacher)] }),
            new TableCell({ children: [new Paragraph(r.student)] }),
            new TableCell({ children: [new Paragraph(r.assessment)] }),
          ],
        })
    ),
  ];

  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ text: (session.collegeName || 'КОЛЛЕДЖ').toUpperCase(), alignment: AlignmentType.CENTER }),
          new Paragraph({ text: 'САБАҚ ЖОСПАРЫ (КСП ТжКБ)', alignment: AlignmentType.CENTER }),
          new Paragraph({ text: `Оқытушы: ${session.fullName || 'Оқытушы'}` }),
          new Paragraph({ text: `Пән/Модуль: ${parsedData.subject || 'Кәсіптік модуль'}` }),
          new Paragraph({ text: `Тақырып: ${parsedData.topic}` }),
          new Paragraph({ text: `Оқу нәтижесі: ${parsedData.objective}` }),
          new Paragraph({ text: '' }),
          new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows }),
        ],
      },
    ],
  });

  const docxBuffer = await Packer.toBuffer(doc);

  // Генерация PDF
  const pdfDoc = new jsPDF();
  pdfDoc.setFontSize(14);
  pdfDoc.text(session.collegeName || 'Колледж', 105, 15, { align: 'center' });
  pdfDoc.setFontSize(11);
  pdfDoc.text(`Оқытушы: ${session.fullName || 'Оқытушы'} | Тақырып: ${parsedData.topic}`, 14, 25);

  autoTable(pdfDoc, {
    startY: 32,
    head: [['Кезең', 'Уақыт', 'Оқытушы', 'Студент', 'Бағалау']],
    body: parsedData.table.map((r) => [r.stage, r.time, r.teacher, r.student, r.assessment]),
    theme: 'grid',
    styles: { fontSize: 8 },
  });

  const pdfBuffer = Buffer.from(pdfDoc.output('arraybuffer'));

  // Отправка файлов в Telegram
  await ctx.reply(
    `✅ **Құжаттар топтамасы сәтті жасалды!**\n\n🏛 Колледж: ${session.collegeName}\n👤 Оқытушы: ${session.fullName}\n📚 Тақырып: ${parsedData.topic}`
  );
  await ctx.replyWithDocument({ source: docxBuffer, filename: `КСП_${parsedData.topic}.docx` });
  await ctx.replyWithDocument({ source: pdfBuffer, filename: `КСП_${parsedData.topic}.pdf` });

  // ОБЯЗАТЕЛЬНОЕ ТРЕБОВАНИЕ: вывод остатка лимита отдельным сообщением
  await ctx.reply(keyPool.getQuotaText());
});

// Запуск бота
if (BOT_TOKEN) {
  bot.launch().then(() => console.log('UstazAI Telegram Bot запущен!'));
}

// Простой веб-сервер, чтобы бесплатный хостинг Render не выключал приложение
const server = http.createServer((req, res) => {
  res.writeHead(200, { 'Content-Type': 'text/plain; charset=utf-8' });
  res.end('UstazAI Telegram Bot белсенді жұмыс істеп тұр!');
});
const PORT = process.env.PORT || 3000;
server.listen(PORT, () => console.log(`Server listening on port ${PORT}`));
