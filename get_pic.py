# this project proposes a method to extract charts from PDF reports
# packages that need to be installed include: pdfminer3k, fitz, PyMuPDF
from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed
from PIL import Image
import fitz
import re
import os


class GetPic:
    def __init__(self, filename, password=''):
        """
        initialization
        :param filename: PDF file path
        :param password: password
        """
        with open(filename, 'rb') as file:
            # create a document analyzer
            self.parser = PDFParser(file)
        # create documentation
        self.doc = PDFDocument()
        # connect documents with document analyzer
        self.parser.set_document(self.doc)
        self.doc.set_parser(self.parser)
        # initialization, get the initial password, or an empty string if none
        self.doc.initialize(password)
        # check whether the document provides txt conversion, ignore it if not, throw an exception
        if not self.doc.is_extractable:
            raise PDFTextExtractionNotAllowed
        else:
            # create a PDF explorer to manage shared resources
            self.resource_manager = PDFResourceManager()
            # create a PDF device object
            self.laparams = LAParams()
            self.device = PDFPageAggregator(self.resource_manager, laparams=self.laparams)
            # create a PDF interpreter object
            self.interpreter = PDFPageInterpreter(self.resource_manager, self.device)
            # create a list of PDF page objects
            self.doc_pdfs = list(self.doc.get_pages())
        # open the PDF file and generate an iterable object containing image doc objects
        self.doc_pics = fitz.open(filename)
        self.pic_info = {}

    def to_pic(self, doc, zoom, pg, pic_path):
        """
        convert single page PDF to pic
        :param doc: the doc object of the image
        :param zoom: image scaling, type int, the larger the value, the higher the resolution
        :param pg: the index of the object in doc_pics
        :param pic_path: image save path
        :return: image path
        """
        rotate = int(0)
        trans = fitz.Matrix(zoom, zoom).preRotate(rotate)
        pm = doc.getPixmap(matrix=trans, alpha=False)
        path = os.path.join(pic_path, str(pg)) + '.png'
        pm.writePNG(path)
        return path
    
    def get_level(self, zip_loc_list):
        """
        record the relative horizontal position of the chart on each page of the PDF
        from top to bottom are level_0, level_1, ..., level_n
        there can be several charts on the same level, i.e. the same level_k(k=0,1,...,n)
        :param zip_loc_list: contains two arrays of loc_top and loc_bottom, and the elements of these two arrays correspond one-to-one
        example:
        [(((33.8, 727.6025, 152.91049999999998, 737.4499999999999), '表 9:微盟-收入利润表'), 
        ((33.250054999999996, 388.634, 198.79753000000002, 396.95000000000005), '资料来源:公司财报,安信证券研究中心预测')), 
        (((33.8, 312.48249999999996, 158.16049999999998, 322.33), '表 10:微盟-现金流量表'), 
        ((33.250054999999996, 123.634, 198.79753000000002, 131.95), '资料来源:公司财报,安信证券研究中心预测'))]
        :return: a dictionary, the key is the relative horizontal position, the value is the corresponding chart list, and the value in the list is the index of the corresponding chart in zip_loc_list
        example:
        {'level_0': [0], 'level_1': [1], 'level_2': [2, 3]}
        """
        level_dict = {}
        length = len(zip_loc_list)
        visit = [0] * length
        level_count = 0
        for i in range(length):
            # a value of 1 indicates that the corresponding chart has already determined its relative horizontal position, skip it directly
            if visit[i] == 1:
                continue
            # the chart corresponding to i is at a different horizontal position than the chart whose position was previously determined
            visit[i] = 1
            level_count_str = 'level_' + str(level_count)
            level_count += 1
            level_dict[level_count_str] = [i]
            # the corresponding charts before i have already determined their relative horizontal positions, so there is no need to traverse them
            for j in range(i + 1, length):
                # a value of 1 indicates that the corresponding chart has already determined its relative horizontal position, skip it directly
                if visit[j] == 1:
                    continue

                # save the height value corresponding to the flatter chart of the two charts to be compared
                high = min(abs(zip_loc_list[i][0][0][3] - zip_loc_list[i][1][0][1]), \
                    abs(zip_loc_list[j][0][0][3] - zip_loc_list[j][1][0][1]))
                
                # if one chart completely surrounds the other, i.e. the upper border is higher than the other, and the lower border is lower than the other, then the two charts are considered to be on a horizontal line
                if zip_loc_list[i][0][0][3] <= zip_loc_list[j][0][0][3] and \
                    zip_loc_list[i][1][0][1] >= zip_loc_list[j][1][0][1]:
                    visit[j] = 1
                    level_dict[level_count_str].append(j)
                
                elif zip_loc_list[i][0][0][3] >= zip_loc_list[j][0][0][3] and \
                    zip_loc_list[i][1][0][1] <= zip_loc_list[j][1][0][1]:
                    visit[j] = 1
                    level_dict[level_count_str].append(j)
                
                # if the difference between the height of the upper border and the height of the lower border of the two charts is greater than the value of high, then the two charts are not considered to be on a horizontal line
                elif min(abs(zip_loc_list[i][0][0][3] - zip_loc_list[j][0][0][3]), \
                    abs(zip_loc_list[i][1][0][1] - zip_loc_list[j][1][0][1])) >= high:
                    continue
                
                # in other cases, the two charts are considered to be on a horizontal line
                else:
                    visit[j] = 1
                    level_dict[level_count_str].append(j)
        return level_dict

    def get_pic_loc(self, doc, pgn):
        """
        get the position of an image in a single page
        :param doc: doc object for PDF
        :return: a tuple, elements include loc_top, loc_bottom, the dimensions of the PDF, the leftmost and rightmost coordinates of all LT objects of a single-page PDF
        """
        self.interpreter.process_page(doc)
        layout = self.device.get_result()
        # the dimensions of the PDF, tuple, (width, height)
        canvas_size = layout.bbox
        # the coordinates of the chart name
        loc_top = []
        # the coordinates of the data source
        loc_bottom = []
        # left and right are used to record the leftmost and rightmost coordinates of all LT objects in a single-page PDF
        left = canvas_size[2]
        right = 0
        # text_order is used to record the first LT object containing text in the current page PDF
        # first_text is used to record its text content
        text_order = 0
        first_text = ''
        # iterate over all LT objects of a single-page PDF
        for i in layout:
            left = min(left, i.bbox[0])
            right = max(right, i.bbox[2])
            if hasattr(i, 'get_text'):
                text = i.get_text().strip()
                text_order += 1
                if text_order == 1:
                    first_text = text
                # match keywords
                if re.search(r'[图表]+\s*\d+[:：\s]*', text):
                    title_start = re.search(r'[图表]+\s*\d+[:：\s]*', text).start()
                    loc_top.append((i.bbox, text[title_start: ].replace('\n', '')))
                elif re.search(r'来源[:：\s]', text):
                    text = text.split('\n')[0]
                    loc_bottom.append((i.bbox, text))
                    # if the first keyword obtained on a single page is r'来源[:：\s]' instead of r'[图表]+\s*\d+[:：\s]*'
                    # then it is possible that this is a cross-page chart, and we need to go to the previous page to find out whether there is corresponding keyword r'[图表]+\s*\d+[:：\s]*'
                    if len(loc_bottom) > 0 and len(loc_top) == 0:
                        last_pgn = pgn - 1
                        # self.pic_info[last_pgn]['loc_bottom']) < len(self.pic_info[last_pgn]['loc_top']
                        # it means that the keyword r'[图表]+\s*\d+[:：\s]*' exists on the previous page and is not matched with the corresponding keyword r'来源[:：\s]'
                        # It can be considered that the keyword r'[图表]+\s*\d+[:：\s]*' on the previous page until the bottom part corresponds to the upper half of the chart across the pages
                        # The current page keyword r'来源[:：\s]' and above until the top part corresponds to the lower part of the spread chart
                        if last_pgn in self.pic_info and \
                            len(self.pic_info[last_pgn]['loc_bottom']) < len(self.pic_info[last_pgn]['loc_top']):
                            # Save the coordinates of the bottom edge of the previous page
                            self.pic_info[last_pgn]['loc_bottom'].append(((0, 0, 0, 0), ''))
                            # Save the coordinates of the top edge of the current page
                            loc_top.append(((0, canvas_size[3], 0, canvas_size[3]), \
                                self.pic_info[last_pgn]['loc_top'][-1][1] + '@~@continue'))
                        # if there is no keyword r'[图表]+\s*\d+[:：\s]*' on the previous page that is paired with the corresponding keyword r'来源[:：\s]' on the current page
                        # then the keyword r'来源[:：\s]' needs to be removed
                        else:
                            if text_order == 1:  
                                loc_top.append(((0, canvas_size[3], 0, canvas_size[3]), first_text))
                            else:
                                loc_bottom.pop(-1)

        return (loc_top, loc_bottom, canvas_size, left, right)
        
    def get_crops(self, pic_path, canvas_size, position, cropped_pic_name, cropped_pic_path):
        """
        extract charts by given coordinates
        :param pic_path: the path of the pic
        :param canvas_size: the size of the original PDF corresponding to the pic, tuple, (0, 0, width, height)
        :param position: the coordinates of the chart to be extracted, tuple, (x1, y1, x2, y2)
        :param cropped_pic_name: the name of the chart to be extracted
        :param cropped_pic_path: the save path of the chart to be extracted
        :return:
        """
        img = Image.open(pic_path)
        # the size of the current image tuple(width, height)
        pic_size = img.size
        # size_increase is a buffer value as margin for extraction
        size_increase = 10
        x1 = max(pic_size[0] * (position[0] - size_increase)/canvas_size[2], 0)
        x2 = min(pic_size[0] * (position[2] + size_increase)/canvas_size[2], pic_size[0])
        y1 = max(pic_size[1] * (1 - (position[3] + size_increase)/canvas_size[3]), 0)
        y2 = min(pic_size[1] * (1 - (position[1] - size_increase)/canvas_size[3]), pic_size[1])
        try:
            cropped_img = img.crop((x1, y1, x2, y2))
            # path to save the extracted chart
            path = os.path.join(cropped_pic_path, cropped_pic_name) + '.png'
            cropped_img.save(path)
            print('we successfully extract:', cropped_pic_name)
        except Exception as e:
            print(e)

    def get_pic_info(self, pic_path, page_count):
        """
        Convert PDF to pictures by page and save related information in dictionary self.pic_info by page
        The key is the page number of the PDF, and the value is the corresponding relevant information
        :param pic_path: the path of the pic
        :param page_count: the number of pages in the PDF
        :return:
        """
        if page_count <= 0:
            return
        for pgn in range(page_count):
            # get the doc of the current page
            doc_pdf = self.doc_pdfs[pgn]
            doc_pic = self.doc_pics[pgn]
            # convert the current page to PNG, the return value is the pic path
            path = self.to_pic(doc_pic, 2, pgn, pic_path)
            pgn_info = self.get_pic_loc(doc_pdf, pgn)
            self.pic_info[pgn] = {
                'path': path,
                'loc_top': pgn_info[0], 
                'loc_bottom': pgn_info[1],
                'canvas_size': pgn_info[2],
                'left': pgn_info[3],
                'right': pgn_info[4]
            }
    
    def generate_result(self, cropped_pic_path):
        """
        extract charts by page from PDF. The charts are extracted from the pic converted from PDF, not directly in PDF
        :param cropped_pic_path: the save path of the chart to be extracted
        :return:
        """
        for k, v in self.pic_info.items():
            loc_list = list(zip(v['loc_top'], v['loc_bottom']))
            # if loc_list is empty, it means that there is no chart in the PDF on this page, skip it directly
            if not loc_list:
                continue
            width = abs(v['right'] - v['left'])
            # reorder the chart list in ascending order by the x1 coordinate of the chart
            loc_list.sort(key=lambda x: x[0][0][0])
            path = v['path']
            canvas_size = v['canvas_size']
            level_dict = self.get_level(loc_list)
            loc_name_pic = []

            for level, item_list in level_dict.items():
                level_count = len(item_list)
                level_order = 0
                # on a horizontal line, if there are several charts, then divide the width into equal parts for calculating the x1 and x2 coordinates
                for item in item_list:
                    x1 = v['left'] + level_order * width / level_count
                    if level_order > 0:
                        x1 = min(min(x1, loc_list[item_list[level_order]][0][0][0]), \
                            loc_list[item_list[level_order]][0][0][2])
                    level_order += 1
                    x2 = v['left'] + level_order * width / level_count
                    # If there are multiple charts on a horizontal line and the current chart is not the rightmost chart
                    # Compare x2 with the x1 coordinate of the chart on the right side of the current chart, and take the smaller value as the x2 coordinate of the current chart
                    if level_count > 1 and level_order < level_count:
                        x2 = min(min(x2, loc_list[item_list[level_order]][0][0][0]), \
                            loc_list[item_list[level_order]][0][0][2])
                    y1 = loc_list[item][1][0][1]
                    y2 = loc_list[item][0][0][3]
                    name = loc_list[item][0][1]
                    loc_name_pic.append((name, (x1, y1, x2, y2)))
                level_order = 0

            if loc_name_pic:
                for i in loc_name_pic:
                    position = i[1]
                    cropped_pic_name = re.sub('/', '_', i[0])
                    self.get_crops(path, canvas_size, position, cropped_pic_name, cropped_pic_path)

    def blend_pic(self, cropped_pic_path):
        """
        traverse the obtained chart, splicing the upper and lower parts of the cross-page chart into one chart and save it, and delete the original upper and lower parts of the chart
        :param cropped_pic_path: the save path of the chart to be extracted
        :return:
        """
        pic_count_dict = {}
        pic_list = os.listdir(cropped_pic_path)
        # Save the picture to the dictionary pic_count_dict according to the name. If it is a cross-page chart, the same name will correspond to multiple pics
        for pic in pic_list:
            pic_name = pic.split('.')[0].split('@~@')[0]
            if pic_name in pic_count_dict:
                pic_count_dict[pic_name].append(pic)
            else:
                pic_count_dict[pic_name] = [pic]
        print(pic_count_dict)

        for k, v in pic_count_dict.items():
            if len(v) == 1:
                continue
            png1 = os.path.join(cropped_pic_path, v[0])
            png2 = os.path.join(cropped_pic_path, v[1])
            # to avoid errors in the upper and lower order, a judgment was added
            if '@~@continue' in png1:
                png1, png2 = png2, png1
            img1, img2 = Image.open(png1), Image.open(png2)
            size1, size2 = img1.size, img2.size
            # the width of the new pic is the smaller of the two pics, and the height is the sum of the heights of the two pics
            joint = Image.new('RGB', (min(size1[0], size2[0]), size1[1] + size2[1]))
            loc1, loc2 = (0, 0), (0, size1[1])
            joint.paste(img1, loc1)
            joint.paste(img2, loc2)
            # delete the original pic and save the new pic
            os.remove(png1)
            os.remove(png2)
            joint.save(png1)


if __name__ == '__main__':
    pdf_path = 'SaaS龙头深耕微信生态，双轮驱动增长.pdf'
    pic_path = 'SaaS龙头深耕微信生态，双轮驱动增长/PNG'
    cropped_pic_path = 'SaaS龙头深耕微信生态，双轮驱动增长/CROPPED_PIC'

    gp = GetPic(pdf_path)
    if not os.path.exists(pic_path):
        os.makedirs(pic_path)
    if not os.path.exists(cropped_pic_path):
        os.makedirs(cropped_pic_path)
    page_count = gp.doc_pics.pageCount
    gp.get_pic_info(pic_path, page_count)
    gp.generate_result(cropped_pic_path)
    gp.blend_pic(cropped_pic_path)