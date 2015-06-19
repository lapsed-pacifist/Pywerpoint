from itertools import product
import numpy as np
import pandas as pd
from win32com.client import pywintypes
from collections import namedtuple

ALIGN_LABELS = 'bottom bottom_base middle top top_base mixed'.split()
ALIGN_LABELS_N = {k: i for i, k in enumerate(ALIGN_LABELS)}

class Win32Interface(object):
	def __init__(self, win32_object):
		super(Win32Interface, self).__setattr__('win32_object', win32_object)

	def __setattr__(self, k, v):
		if k in self.properties:
			super(Win32Interface, self).__setattr__(k, v)
		else:
			setattr(self.win32_object, k, v)

	def __getattr__(self, v):
		return getattr(self.win32_object, v)

	@property
	def properties(self):
		class_items = self.__class__.__dict__.iteritems()
		return {k:v for k, v in class_items if isinstance(v, property) and k != 'properties'}


class Table(Win32Interface):
	def _win32_index(self, i, axis='row'):
		if axis == 'row':
			l = self.win32_object.Table.Rows.Count
		elif axis == 'column':
			l = self.win32_object.Table.Columns.Count
		else:
			raise ValueError('{} not an axis'.format(axis))
		return l + 1 if i is None else i + 1 if i >= 0 else l + i + 1

	def _get_index(self, key):
		if isinstance(key, tuple):
			key = list(key)
		else:
			key = [key, slice(None)]
		for i, (k, count, axis) in enumerate(zip(key, self.shape, ['row', 'column'])):
			if isinstance(k, slice):
				start = 0 if k.start is None else k.start
				stop = count if k.stop is None else k.stop
				step = 1 if k.step is None else k.step
				start, stop = map(lambda x: self._win32_index(x, axis), [start, stop])
				key[i] = range(start, stop, step)
			else:
				key[i] = [self._win32_index(k, axis)]
		return key

	def __getitem__(self, key):
		key = self._get_index(key)
		map_cell = lambda coord: self.win32_object.Table.Cell(coord[0], coord[1])
		mapped = map(map_cell, list(product(*key)))
		mapped = map(lambda x: Cell(x), mapped)
		mapped = np.array(mapped).reshape(map(len, key))
		if mapped.shape == (1,1):
			return mapped[0, 0]
		return CellRange(mapped)

	def __setitem__(self, key, value):
		raise NotImplementedError("You must select what attribute of the table to set.\neg.\n"\
			">>> table[1,1].text = 'some text'\nOR\n>>> table[1,1].font = 'a font'")

	def adjust_size(self, height=399.1181102362):
		self.height = height
		self.width =  960 - (2*left)

	def adjust(self, height=399.1181102362, left=36.28346456693, top=113.1023622047):
		self.height = height
		self.width =  960 - (2*left)
		self.left = left
		self.top = top

	def scale(self, percent):
		self.Table.ScaleProportionally(percent)

	def trim_text(self):
		self[:,:]._apply_map(lambda x: x.trim_text())

	# def fit_columns(self, width_percentages=None):
	# 	width_percentages = [0]*len(self.columns)
	# 	assert len(self.columns) == len(width_percentages)
	# 	for w, c in zip(width_percentages, self.columns):
	# 		c.width = w

	# def minimize_table(self):
	# 	self.fit_columns()
	# 	self.adjust()

	# def one_line_rows(self, irows):
	# 	self.

	@property
	def columns(self):
		return Axis(self.Table.Columns)

	@property
	def rows(self):
		return Axis(self.Table.Rows)

	@property
	def cells(self):
		return self[:,:]

	@property
	def shape(self):
		return tuple(getattr(self.win32_object.Table, i).Count for i in ['Rows', 'Columns'])

	@property
	def borders(self):
		return Borders(self.cells)

	def __repr__(self):
		return self.cells.__repr__()

	def insert(self, index, axis):
		if axis == 'column':
			assert index <= self.shape[1]
			index = self._win32_index(index, 'column')
			self.Table.Columns.Add(index)
		elif axis == 'row':
			assert index <= self.shape[0]
			index = self._win32_index(index, 'row')
			self.Table.Rows.Add(index)
		else:
			raise ValueError('axis has to be "column" or "row"')

	def delete(self, index, axis):
		if axis == 'column':
			assert index <= self.shape[1]
			index = self._win32_index(index, 'column')
			self.Table.Columns(index).Delete()
		elif axis == 'row':
			assert index <= self.shape[0]
			index = self._win32_index(index, 'row')
			self.Table.Rows(index).Delete()
		else:
			raise ValueError('axis has to be "column" or "row"')

	def __delitem__(self, key):
		key = self._get_index(key)
		if all(len(k) == l for k, l in zip(key, self.shape)):
			self.Delete()
		elif len(key[0]) == self.shape[0]:
			for k in sorted(key[1], reverse=True):
				self.delete(k-1, 'column')
		elif len(key[1]) == self.shape[1]:
			for k in sorted(key[0], reverse=True):
				self.delete(k-1, 'row')
		else:
			raise ValueError("Cannot delete a subset of cells, only rows/columns!")


class Cell(Win32Interface):
	def __repr__(self):
		return self.Shape.TextFrame.TextRange.Text

	@property
	def parent_table(self):
		return Table(self.win32_object.Parent.Parent)

	@property
	def text(self):
		return self.Shape.TextFrame.TextRange.Text

	@text.setter
	def text(self, v):
		self.Shape.TextFrame.TextRange.Text = v

	@property
	def font(self):
		return self.Shape.TextFrame.TextRange.Font.Name

	@font.setter
	def font(self, v):
		self.Shape.TextFrame.TextRange.Font.Name = v

	@property
	def text_colour(self):
		return self.Shape.TextFrame.TextRange.Font.Color

	@text_colour.setter
	def text_colour(self, v):
		self.Shape.TextFrame.TextRange.Font.Color = v

	@property
	def fill_colour(self):
		return self.Shape.Fill.ForeColor.RGB

	@fill_colour.setter
	def fill_colour(self, v):
		self.Shape.Fill.ForeColor.RGB = v

	@property
	def text_size(self):
		return self.Shape.TextFrame.TextRange.Font.Size

	@text_size.setter
	def text_size(self, v):
		self.Shape.TextFrame.TextRange.Font.Size = v

	@property
	def edge_colour(self):
		return [b.colour for b in self.borders]

	@edge_colour.setter
	def edge_colour(self, v):
		for b in self.borders:
			b.colour = v

	@property
	def edge_weight(self):
		return [b.colour for b in self.borders]

	@edge_weight.setter
	def edge_weight(self, v):
		for b in self.borders:
			b.Weight = v

	@property
	def borders(self):
		return Borders(self)

	@property
	def hyperlink(self):
		return self.Shape.TextFrame.TextRange.TextRange.ActionSettings(1).Hyperlink.Address

	@hyperlink.setter
	def hyperlink(self, v):
		self.Shape.TextFrame.TextRange.ActionSettings(1).Action = 7
		self.Shape.TextFrame.TextRange.ActionSettings(1).Hyperlink.Address = v

	def trim_text(self):
		x = self.Shape.TextFrame.TextRange
		self.text = x.Lines(1).Text[:-3] + '...' if len(x.Lines()) > 1 else x.Text


class Border(Win32Interface):
	@property
	def colour(self):
		return self.ForeColor.RGB

	@colour.setter
	def colour(self, v):
		self.ForeColor.RGB = v


class Borders(object):
	"""contains the borders for a Cell or CellRange. Able to access/apply arraywise"""
	def __init__(self, cell_object):
		assert isinstance(cell_object, Cell) or isinstance(cell_object, CellRange)
		super(Borders, self).__init__()
		super(Borders, self).__setattr__('cell_object', cell_object)
		super(Borders, self).__setattr__('_labels', 'bottom left right top'.split())
		self.make_borders()

	def _apply_map(self, f):
		func = np.vectorize(f)
		return func(self.cell_object)

	def make_borders(self):
		for l, i in zip(self._labels, range(1,5)):
			try:
				bs = CellRange(self.cell_object._apply_map(lambda x: Border(x.Borders(i))))
			except AttributeError:
				bs = Border(self.cell_object.Borders(i))
			super(Borders, self).__setattr__(l, bs)

	def __getattr__(self, k):
		return {i: getattr(getattr(self, i), k) for i in self._labels}

	def __setattr__(self, k, v):
		if k not in self.properties:
			for l in self._labels:
				setattr(getattr(self, l), k, v)
		else:
			super(Borders, self).__setattr__(k, v)
		
	@property
	def properties(self):
		class_items = self.__class__.__dict__.iteritems()
		return {k:v for k, v in class_items if isinstance(v, property) and k != 'properties'}


class CellRange(object):
	def __init__(self, cell_array):
		super(CellRange, self).__init__()
		super(CellRange, self).__setattr__('cell_array', cell_array)

	def _apply_map(self, f):
		func = np.vectorize(f)
		return func(self.cell_array)

	def __getattr__(self, k):
		try:
			return getattr(self.cell_array, k)
		except (AttributeError, pywintypes.com_error):
			arr = self._apply_map(lambda x: getattr(x, k))
			return CellRange(arr)

	def __setattr__(self, k, v):
		if k in self.properties:
			super(CellRange, self).__setattr__(k, v)
		else:
			if hasattr(v, 'shape'):
				if isinstance(v, np.ndarray):
					assert self.shape == v.shape, 'mismatched shape'
					v_unravelled = v.ravel()
				elif isinstance(v, pd.core.frame.DataFrame):
					assert (v.shape[0]+1, v.shape[1]) == self.shape, 'mismatched shape'
					v_array = np.empty([v.shape[0]+1, v.shape[1]], dtype=object)
					v_array[1:, :] = v.values
					v_array[0, :] = v.columns.tolist()
					v_unravelled = v_array.ravel()
				else:
					raise TypeError("Strange. This object has a shape but is not a DataFrame/numpy array. Suspicious...")
				for cell, value in zip(self.cell_array.ravel(), v_unravelled):
					cell.__setattr__(k, value)
			else:
				self._apply_map(lambda x: setattr(x, k, v))

	def __getitem__(self, index):
		return CellRange(self.cell_array[index])
	
	def __repr__(self):
		return self.cell_array.__repr__()

	@property
	def parent_table(self):
		return Table(self[0,0].Parent.Parent)

	@property
	def properties(self):
		class_items = self.__class__.__dict__.iteritems()
		return {k:v for k, v in class_items if isinstance(v, property) and k != 'properties'}


class Axis(Win32Interface):
	def __len__(self):
		return self.win32_object.Count

	def _win32_index(self, i):
		l = self.__len__()
		return l + 1 if i is None else i + 1 if i >= 0 else l + i + 1

	def __getitem__(self, v):
		ind = self._win32_index(v)
		if ind > self.__len__():
			raise IndexError
		return self.win32_object(ind)


class Image(Win32Interface):
	@property
	def height(self):
		return self.TextFrame.parent.height

	@height.setter
	def height(self, v):
		self.TextFrame.parent.height = v

	@property
	def width(self):
		return self.TextFrame.parent.width

	@width.setter
	def width(self, v):
		self.TextFrame.parent.width = v

	@property
	def left(self):
		return self.TextFrame.parent.left

	@left.setter
	def left(self, v):
		self.TextFrame.parent.left = v

	@property
	def top(self):
		return self.TextFrame.parent.top

	@top.setter
	def top(self, v):
		self.TextFrame.parent.top = v

	@property
	def position(self):
		return (self.TextFrame.parent.left, self.TextFrame.parent.top)

	@position.setter
	def position(self, v):
		self.TextFrame.parent.left, self.TextFrame.parent.top = v

	@property
	def size(self):
		return (self.TextFrame.parent.height, self.TextFrame.parent.width)

	@size.setter
	def size(self, v):
		self.TextFrame.parent.height, self.TextFrame.parent.width = v	

	def absolute_resize(self, height=None, width=None, keep_aspect_ratio=True):
		if width is None and height is None:
			width, height = self.width, self.height
		elif width is not None and height is not None:
			if keep_aspect_ratio:
				raise ValueError("cannot keep aspect ratio if width and height are both specified")
		elif width is None:
			if keep_aspect_ratio:
				width = self.width * height * 1. / self.height
			else:
				width = self.width
		elif height is None:
			if keep_aspect_ratio:
				height = self.height * width * 1. / self.width
			else:
				height = self.height
		self.height = height
		self.width = width

	def relative_resize(self, height=None, width=None, keep_aspect_ratio=True):
		if width is None and height is None:
			width, height = 1, 1
		elif width is not None and height is not None:
			if keep_aspect_ratio:
				raise ValueError("cannot keep aspect ratio if width and height are both specified")
		elif width is None:
			if keep_aspect_ratio:
				width = height
			else:
				width = 1
		elif height is None:
			if keep_aspect_ratio:
				height = width
			else:
				height = 1
		assert 0 < height <= 1, "a relative height must be above 0 and below or equal to 1"
		assert 0 < width <= 1, "a relative width must be above 0 and below or equal to 1"
		self.height *= height
		self.width *= width


class TextBox(Win32Interface):
	pass

