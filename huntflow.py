#!/usr/bin/env python
# -*- coding: utf-8 -*-
__author__ = "Nikita Markozov"

import sys
import os
import re

import xlrd
import requests


def upload_file(name: str, path: str, url_: str, headers: str) -> dict:
	"""
	Функция для загрузки резюме в виде файла
	"""
	url = url_ + 'upload'
	files = {'file': (f'{name}.pdf', open(f'{path}', 'rb'), 'application/pdf')}
	req_resume = requests.post(url, headers=headers, files=files).json()
	return req_resume


def create_body_to_req(last_name, first_name, middle_name, position, money,
						birth_day, month, year, photo, text, resume_id) -> dict:
	"""
	Сборщик тела для создания кандидата
	"""
	
	body = {
		"last_name": last_name,
		"first_name": first_name,
		"middle_name": middle_name,
		"position": position,
		"money": money,
		"birthday_day": birth_day,
    	"birthday_month": month,
    	"birthday_year": year,
    	"photo": photo,
		"externals": [
		{
			"data":{
				"body": text,
			},
			"auth_type": "NATIVE",
			"files": [{
				"id": resume_id
			}]
		}
		]
	}
	return body


def create_body_to_mark(id_: int, text: str, status_id: int, vacancy_id: int) -> dict:
	"""
	Сборщик тела для запроса с комментариями к резюме в вакансии
	"""

	body = {
		"vacancy": vacancy_id,
		"status": status_id,
		"comment": text,
		"files": [{
			"id": id_,
		}]
	}
	return body


def request_to_add_applicant(info_list: list, resume: dict, url_: str, headers: dict) -> int:
	"""
	Функция для добавления кандидата
	"""
	name_info = str(info_list[1].value)
	last_name = name_info.split()[0]
	first_name = name_info.split()[1]
	position = info_list[0].value
	money = str(info_list[2].value)
	try:
		middle_name = name_info.split()[2]
	except IndexError as e:
		middle_name = None
	photo = resume['photo']['id']
	text = resume['text']
	resume_id = resume['id']
	try:
		birth = resume['fields']['birthdate']
		birth_day = birth['day']
		month = birth['month']
		year = birth['year']

	except (KeyError, TypeError) as e:
		birth_day = None
		month = None
		year = None

	url = url_ + 'applicants'

	body = create_body_to_req(last_name, first_name, middle_name, position, money,
							birth_day, month, year, photo, text, resume_id)

	json_adding_applicant = requests.post(url, json=body, headers=headers).json()
	return json_adding_applicant['id']
	

def mark_request(list_: list, id_list: list, url_: str, headers: dict) -> None:
	"""
	Функция для поиска кандидата и добавление в работу с вакансией + добавление комментария
	"""

	req = requests.get(url_ + 'vacancies', headers=headers).json()
	vacancie_position = [{i['position']: i['id']} for i in req['items']]
	positions = []
	bodys = []
	for i in range(len(id_list)):
		text_comment = list_[i][-1].value
		if text_comment.lower() != 'отказ':
			status_id = 3
		else:
			status_id = 10
		request = requests.get(url_ + 'applicants/' + str(id_list[i]), headers=headers).json()
		pos_from_json = request['position']
		positions.append(pos_from_json)
		for j in vacancie_position:
			if pos_from_json in j.keys():
				vacancy_id = j[pos_from_json]
		body = create_body_to_mark(id_list[i], text_comment, status_id, vacancy_id)
		bodys.append(body)
			
	for i in vacancie_position:
		for j in range(0, len(positions)):
			if positions[j] in i.keys():
				
				request_to_add = requests.post(url_ + 'applicants/' + str(id_list[j]) + '/vacancy', 
				 							json=bodys[j], headers=headers).json()
	

def find_resume_files(path: str, folder: str, name: str) -> str:
	"""
	Функция для поиска файла с резюме по вложенным папкам
	"""

	local_path = path + '/' + str(folder.value)
	for i in os.listdir(local_path):
		splited_name_value = str(name.value).split()[0]
		if re.search(splited_name_value, i):
			file_path = local_path + '/' + i
			return file_path



def xlsx(path: str) -> list:
	"""
	Функция для прочитки excel, возвращает список со всеми колонками и информацией в них
	"""

	files = os.listdir(path)
	xls = [i for i in files if re.search('.xlsx', i)]
	xl_book_path = re.sub(' ', '', path + '/' + xls[0])
	book = xlrd.open_workbook(xl_book_path)
	sh = book.sheet_by_index(0)
	row_with_info = [sh.row(rx) for rx in range(1, sh.nrows)]
	return row_with_info


def main(path: str, token: str) -> str:
	"""
	Главная функция, которая запускает все остальные
	"""
	headers = {
		'User-Agent': 'App/1.0 (drozdivoron@gmail.com)',
    	'Authorization': f'Bearer {token}',
	    'X-File-Parse': 'true',
	}
	url = 'https://dev-100-api.huntflow.ru/account/2/'

	infos = xlsx(path)
	list_with_resume = [find_resume_files(path, i[0], i[1]) for i in infos]
	for_itter = range(len(infos))
	resumes_json = [upload_file(infos[i][1], list_with_resume[i], url, headers) for i in for_itter]
	applicants_id = [request_to_add_applicant(infos[i], resumes_json[i], url, headers) for i in for_itter]
	mark_request(infos, applicants_id, url, headers)
	return 'Done'


if __name__ == '__main__':
	refactor_path = sys.argv[1].replace('\\', '/')
	token = str(sys.argv[2])
	if os.listdir(refactor_path):
		try:
			main(rf'{sys.argv[1]}', token)
		except TypeError as e:
			print(e)
		except IndexError:
			print('Отутствует один из аргументов или аргументы не переданы!!')
	else:
		print('Папки не существует')