import telebot
bot = telebot.TeleBot('6102484745:AAF5CQy-pfbm8UaNw92ST6JaphJ9xwcbbnY')

@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, 'Hello')


bot.polling(none_stop=True)
