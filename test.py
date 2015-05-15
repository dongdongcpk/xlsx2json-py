#coding=utf-8

import unittest
import datetime
from xlsx2json import parse_cell_value as pcv

class TestPcv(unittest.TestCase):
	
	def test_int(self):
		a = 18
		self.assertTrue(isinstance(pcv(a), int))

	def test_float(self):
		a = 3.14
		self.assertTrue(isinstance(pcv(a), float))

	def test_date(self):
		a = datetime.datetime.now()
		self.assertEquals(pcv(a), a.ctime())

	def test_bool(self):
		a = True
		self.assertTrue(isinstance(pcv(a), bool))
		b = 'true'
		self.assertTrue(isinstance(pcv(b), bool))
		c = 'FALSE'
		self.assertTrue(isinstance(pcv(c), bool))

	def test_list(self):
		a = u'1,2,3,4'
		self.assertTrue(isinstance(pcv(a), list))
		b = u'1,'
		self.assertTrue(isinstance(pcv(b), list))
		c = u'hello,world'
		parsed_c = pcv(c)
		self.assertTrue(isinstance(parsed_c, list))
		self.assertEquals(parsed_c[0], 'hello')
		d = u'true,false'
		self.assertEquals(pcv(d)[0], True)

	def test_object(self):
		a = u'name:dong,age:18,isMax:true'
		parsed_a = pcv(a)
		self.assertTrue(isinstance(parsed_a, dict))
		self.assertEquals(parsed_a[u'name'], 'dong')
		self.assertEquals(parsed_a[u'age'], 18)
		self.assertEquals(parsed_a[u'isMax'], True)

	def test_string(self):
		a = u'foo'
		self.assertTrue(isinstance(pcv(a), str))
		b = u'你好'
		self.assertTrue(isinstance(pcv(b), str))

if __name__ == '__main__':
	unittest.main()