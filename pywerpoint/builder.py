from __future__ import unicode_literals
import pandas as pd
import os
import numpy as np
import win32com.client as wi
import zipfile
import re
from tables import *
from PIL import Image as PILImage
from matplotlib.colors import hex2color, rgb2hex

layout_table = pd.read_csv("pywerpoint\\slide_layout_enumeration.csv") 
SLIDE_LAYOUTS = {row.Name: int(row.N) for i, row in layout_table.iterrows()}
SLIDE_LAYOUT_NS = {v: k for k, v in SLIDE_LAYOUTS.iteritems()}

def rgb_int2tuple(rgbint):
    return (rgbint // 256 // 256 % 256, rgbint // 256 % 256, rgbint % 256)

def rgb_tuple2int(*rgb):
	try:
		r, g, b = rgb
	except (ValueError, TypeError):
		r, g, b = rgb[0]
	return r + g*256 + b*(256**2)


class Presentation(object):
	def __init__(self, template='table_temp.pptx', close=True):
		super(Presentation, self).__init__()
		self.template = template
		self.close = close
		self.temp_path = os.path.abspath(self.template)
		self.Application = wi.Dispatch("PowerPoint.Application")
		if os.path.exists(self.temp_path):
			self._win32Presentation = self.Application.Presentations.Open(self.temp_path)
		else:
			self._win32Presentation = self.Application.Presentations.Add()

	def __enter__(self):
		return self
	
	def __exit__(self, type, value, traceback):
		if self.close:
			self._win32Presentation.Close()
		# print value[2][2], 
		# return True
		# return isinstance(value, TypeError) #suppresses typeerrors
	
	def save(self):
		self.Application.Presentation.SaveAs(self.temp_path)
	
	def new_slide(self, slide_type, index=None):
		assert index <= self.__len__(), '{} is beyond the Presentation'.format(index)
		try:
			num = SLIDE_LAYOUTS[slide_type]
		except KeyError:
			raise KeyError('{} is not a valid slide type.\nValid ones are:\n{}'.format(slide_type, ','.join(SLIDE_LAYOUTS.keys())))
		i = self._win32_index(index)
		return Slide(self, self._win32Presentation.Slides.Add(i, num))

	def create_slide(self, slide_type, index=None):
		try:
			self.new_slide(slide_type, index)
		except KeyError:
			try:
				self.template_slides[slide_type].copy()
				self.paste_slide(index)
			except KeyError:
				raise KeyError("{} is not a valid slide type or detected template name".format(slide_type))
	
	def paste_slide(self, index):
		assert index <= self.__len__(), "Cannot insert copied slide beyond Presentation"
		i = self._win32_index(index)
		self._win32Presentation.Slides.Paste(i)

	def _win32_index(self, i):
		l = self.__len__()
		return l + 1 if i is None else i + 1 if i >= 0 else l + i + 1

	def _win32_get_item(self, i):
		try:
			_win32_slide = self._win32Presentation.Slides(self._win32_index(i))
			return Slide(self, _win32_slide)
		except wi.pywintypes.com_error:
			raise IndexError
	
	def __getitem__(self, index):
		if isinstance(index, slice):
			start = 0 if index.start is None else index.start
			stop = self.__len__() if index.stop is None else index.stop
			step = 1 if index.step is None else index.step
			return [self._win32_get_item(i) for i in xrange(start, stop, step)]
		return self._win32_get_item(index)

	def __setitem__(self, index, value):
		if isinstance(value, Slide):
			value.copy()
			self.paste_slide(index)
		elif isinstance(value, basestring):
			try:
				self.create_slide(value, index)
			except KeyError:
				self.template_slides[value].copy()
				self.paste_slide(index)
		else:
			raise TypeError('value to be set must be a pre-existing slide or a name of a layout/slide template')
		
	def __len__(self):
		return self._win32Presentation.Slides.Count

	def __delitem__(self, index):
		self._win32Presentation.Slides(self._win32_index(index)).Delete()

	@property
	def template_slides(self):
		res = {}
		for n, slide in enumerate(self):
			txts = [re.match(r'<template: ([^>]+)>', txt.TextFrame.TextRange.Text) for txt in slide.textboxes]
			txt = [i.groups()[0] for i in txts if i is not None][0]
			res[txt] = slide
		return res


class Slide(object):
	def __init__(self, parent_presentation, _win32_slide):
		super(Slide, self).__init__()
		self.parent_presentation = parent_presentation
		self._win32_slide = _win32_slide

	def select(self):
		self._win32_slide.Select()

	def copy(self):
		self._win32_slide.Copy()
		return self._win32_slide

	def __eq__(self, other):
		return self._win32_slide.SlideID == other._win32_slide.SlideID

	def __ne__(self, other):
		return not self.__eq__(other)

	def create_table(self, nrows, ncols, left=None, top=None, height=None, width=None):
		_win32_tab = self._win32_slide.Shapes.AddTable(nrows, ncols, left, top, height, width)
		return Table(self, _win32_tab)

	def insert_image(self, im_dir, left, top, width=None, height=None, keep_aspect_ratio=True, compress=False):
		im_width, im_height = PILImage.open(im_dir).size
		if width is None and height is None:
			width, height = im_width, im_height
		elif width is not None and height is not None:
			if keep_aspect_ratio:
				raise ValueError("cannot keep aspect ratio if width and height are both specified")
		elif width is None:
			if keep_aspect_ratio:
				width = im_width * height * 1. / im_height
			else:
				width = im_width
		elif height is None:
			if keep_aspect_ratio:
				height = im_height * width * 1. / im_width
			else:
				height = im_height

		im = self._win32_slide.Shapes.AddPicture2(FileName=im_dir, LinkToFile=False, SaveWithDocument=True, Left=left, Top=top, Width=width, Height=height, compress=compress)
		return Image(im)

	@property
	def images(self):
		i_list = [i for i in self._win32_slide.Shapes if i.Type in (13,)]
		return map(lambda x: Image(x), i_list)
		 
	@property
	def layout(self):
		return SLIDE_LAYOUT_NS[p[0]._win32_slide.Layout]

	@property
	def index(self):
		return self._win32_slide.SlideNumber - 1

	@property
	def tables(self):
		t_list = [i for i in self._win32_slide.Shapes if hasattr(i, 'table')]
		return map(lambda x: Table(x), t_list)

	@property
	def textboxes(self):
		t_list = [i for i in self._win32_slide.Shapes if hasattr(i, 'TextFrame')]
		res = []
		for t in t_list:
			try:
				t.TextFrame.TextRange.Text
				res.append(t)
			except wi.pywintypes.com_error:
				pass
		return map(lambda x: TextBox(x), res)


if __name__ == '__main__':
	T = pd.DataFrame(np.arange(45).reshape(9,5), columns=list('ABCDE'))
	with Presentation(close=True) as P:
		table = P[0].tables[0]
		print table.cells.text
		table.cells.text = T