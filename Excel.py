import openpyxl
from PIL import Image, ImageDraw, ImageFont

table = openpyxl.open('Test_t.xlsx', read_only=True)
sheet = table.active


font_1 = ImageFont.truetype('Noah-Bold.ttf', size=(14*7))
font_2 = ImageFont.truetype('Noah-Regular.ttf', size=(15*7))

sample = Image.open('Montazhnaya_oblast_2_1.jpg')


cl = str(input('Введите класс\nПример(10m2):'))
if cl == '10m2':
    pr = 66


elif cl == '10m1':
    pr = 64




try:
    for s, x in zip(range(0, 2852, 2851), range(pr + 2, pr, -1)):
        for k in range(710, 1216, 505):
            for i, g in zip(range(1761, 6845, 848), range(2, 28, 5)):

                if ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'ф':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Физика', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'лф':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Лекция физика', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'и':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'История', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'а':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Алгебра', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'ла':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Лекция алгебры', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'г':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Геометрия', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'г':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Геометрия', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'лг':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Лекция геометрии', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'и':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Информатика', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'о':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Обж', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'об':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Обществознание', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'р':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Русский язык', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'л':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Литература', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'ан':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Английский язык', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'х':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Химия', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'фи':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Физкультура', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'пд':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Проэктная д.', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g}'].value)).split(' ')[:-1]) == 'None':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), '', font=font_2, fill=('#000000'))

        for k, h in zip(range(1564, 2169, 604), range(0, 2, 1)):
            for i, g in zip(range(1761, 6845, 848), range(3, 29, 5)):
                if ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'ф':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Физика', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'лф':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Лекция физика', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'и':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'История', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'а':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Алгебра', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'ла':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Лекция алгебры', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'г':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Геометрия', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'г':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Геометрия', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'лг':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Лекция геометрии', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'ин':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Информатика', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'о':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Обж', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'об':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Обществознание', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'р':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Русский язык', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'л':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Литература', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'ан':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Английский язык', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'х':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Химия', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'фи':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Физкультура', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'пд':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k + s), 'Проэктная д.', font=font_2, fill=('#000000'))

                elif ' '.join((str(sheet[f'{chr(x)}{g + h}'].value)).split(' ')[:-1]) == 'None':
                    draw_text_1 = ImageDraw.Draw(sample)
                    draw_text_1.text((i, k), '', font=font_2, fill=('#000000'))
        for l, c in zip(range(5, 31, 5), range(1761, 6845, 848)):
            if ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'ф':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'Физика', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'лф':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'Лекция физика', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'и':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'История', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'а':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'Алгебра', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'ла':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'Лекция алгебры', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'г':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'Геометрия', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'г':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'Геометрия', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'лг':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'Лекция геометрии', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'ин':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'Информатика', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'о':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'Обж', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'об':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'Обществознание', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'р':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'Русский язык', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'л':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'Литература', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'ан':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'Английский язык', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'х':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'Химия', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'фи':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'Физкультура', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'пд':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), 'Проэктная д.', font=font_2, fill=('#000000'))

            elif ' '.join((str(sheet[f'{chr(x)}{l}'].value)).split(' ')[:-1]) == 'None':
                draw_text_1 = ImageDraw.Draw(sample)
                draw_text_1.text((c, 2556 + s), '', font=font_2, fill=('#000000'))

        for m in range(828, 1335, 506):
            for n, v in zip(range(1761, 6845, 848), range(2, 28, 5)):
                if ' '.join((str(sheet[f'{chr(x)}{v}'].value)).split(' ')[-1:]) == 'None':
                    draw_text_2 = ImageDraw.Draw(sample)
                    draw_text_2.text((n, m + s), '', font=font_1, fill=('#727272'))

                else:
                    draw_text_2 = ImageDraw.Draw(sample)
                    draw_text_2.text((n, m + s), ' '.join((str(sheet[f'{chr(x)}{v}'].value)).split(' ')[-1:]), font=font_1, fill=('#727272'))

        for z, q in zip(range(1682, 2287, 604), range(0, 2, 1)):
            for n, v in zip(range(1761, 6845, 848), range(3, 29, 5)):
                if ' '.join((str(sheet[f'{chr(x)}{v+q}'].value)).split(' ')[-1:]) == 'None':
                    draw_text_2 = ImageDraw.Draw(sample)
                    draw_text_2.text((n, z + s), '', font=font_1, fill=('#727272'))

                else:
                    draw_text_2 = ImageDraw.Draw(sample)
                    draw_text_2.text((n, z + s), ' '.join((str(sheet[f'{chr(x)}{v+q}'].value)).split(' ')[-1:]), font=font_1,
                                     fill=('#727272'))

        for b, u in zip(range(1761, 6845, 848), range(5, 31, 5)):
                if ' '.join((str(sheet[f'{chr(x)}{u}'].value)).split(' ')[-1:]) == 'None':
                    draw_text_2 = ImageDraw.Draw(sample)
                    draw_text_2.text((b, 2674 + s), '', font=font_1, fill=('#727272'))

                else:
                    draw_text_2 = ImageDraw.Draw(sample)
                    draw_text_2.text((b, 2674 + s), ' '.join((str(sheet[f'{chr(x)}{u}'].value)).split(' ')[-1:]), font=font_1, fill=('#727272'))


except Exception as ex:
    print(f'Ошибка: {ex}')
way = f'D:\Class_{cl}.jpg'

sample.show() #- эта команда для вывода на экран готового расписания
# sample.save(way) #-эта для сохранения
print(f'Если ты читаешь эту фразу из консоли,\n то значит у тебя все запустилось и я несомненно этому рад.\n Ты создал расписание для: {cl}. \n Оно находится в папке:  {way}')
