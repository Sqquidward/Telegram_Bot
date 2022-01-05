import time
import openpyxl
import datetime
import logging
import aiogram.utils.markdown as md
from aiogram import Bot, Dispatcher, types, executor
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters import Text
from aiogram.dispatcher.filters.state import StatesGroup, State
from aiogram.utils import executor
import sqlite3
import Markup as nav
from Token import API_TOKEN

book = openpyxl.open("Raspisanie.xlsx", read_only=True)
sheet = book.worksheets[1]


base = sqlite3.connect('group.db')
cursor = base.cursor()

base.execute('CREATE TABLE IF NOT EXISTS data(login PRIMARY KEY, class text)')
base.commit()




logging.basicConfig(level=logging.INFO)
API_TOKEN = API_TOKEN

bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot, storage=MemoryStorage())

class Form(StatesGroup):
    classs = State()

@dp.message_handler(commands='start')
async def cmd_start(message: types.Message):
    await Form.classs.set()
    await message.reply("Привет в каком классе ты учишься?\nПример ввода: 10м1")


@dp.message_handler(state=Form.classs)
async def process_name(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['class'] = message.text

        info = cursor.execute('SELECT * FROM data WHERE login == ?', (message.from_user.id,))
        if info.fetchone() is None:
            cursor.execute('INSERT INTO data VALUES(?, ?)', (message.from_user.id, data['class']))
            base.commit()
            await bot.send_message(
                message.chat.id, "Отлично,\nтеперь ты можешь воспользоваться моими функциями", reply_markup=nav.mainMenu
            )

        else:
            cursor.execute('UPDATE data SET class == ? WHERE login == ?', (data['class'], message.from_user.id))
            base.commit()
            await bot.send_message(
                message.chat.id, "Отлично,\nтеперь ты можешь воспользоваться моими функциями", reply_markup=nav.mainMenu
            )

    await state.finish()

@dp.callback_query_handler(lambda c: c.data == 'Rasp_thisday')
async def process_callback_button1(callback_query: types.CallbackQuery):
    await bot.answer_callback_query(callback_query.id)
    await bot.delete_message(callback_query.from_user.id, callback_query.message.message_id)
    await bot.send_message(callback_query.from_user.id, '')

@dp.callback_query_handler(lambda c: c.data == 'raspisanie')
async def process_callback_button1(callback_query: types.CallbackQuery):
    await bot.answer_callback_query(callback_query.id)
    await bot.delete_message(callback_query.from_user.id, callback_query.message.message_id)
    await bot.send_message(callback_query.from_user.id, 'Выберите', reply_markup=nav.Raspisanie)

@dp.callback_query_handler(lambda c: c.data == 'cancel')
async def process_callback_button1(callback_query: types.CallbackQuery):
    await bot.answer_callback_query(callback_query.id)
    await bot.delete_message(callback_query.from_user.id, callback_query.message.message_id)
    await bot.send_message(callback_query.from_user.id, 'Ты можешь воспользоваться моими функциями', reply_markup=nav.mainMenu)

@dp.message_handler(commands=['info'])
async def send_welcome(message: types.Message):
    await message.answer("Данный бот создан специально\nдля учеников лицея, чтобы выполнять нетрудные задачи")
    print(message.from_user.id)

@dp.message_handler(commands=['help'])
async def send_welcome(message: types.Message):
    await message.answer(message.from_user.id)



if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)