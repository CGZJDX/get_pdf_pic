# 需要安装的包包括：pdfminer3k, fitz, PyMuPDF
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
        初始化
        :param filename: pdf路径
        :param password: 密码
        """
        with open(filename, 'rb') as file:
            # 创建文档分析器
            self.parser = PDFParser(file)
        # 创建文档
        self.doc = PDFDocument()
        # 连接文档与文档分析器
        self.parser.set_document(self.doc)
        self.doc.set_parser(self.parser)
        # 初始化, 提供初始密码, 若无则为空字符串
        self.doc.initialize(password)
        # 检测文档是否提供txt转换, 不提供就忽略, 抛出异常
        if not self.doc.is_extractable:
            raise PDFTextExtractionNotAllowed
        else:
            # 创建PDF资源管理器, 管理共享资源
            self.resource_manager = PDFResourceManager()
            # 创建一个PDF设备对象
            self.laparams = LAParams()
            self.device = PDFPageAggregator(self.resource_manager, laparams=self.laparams)
            # 创建一个PDF解释器对象
            self.interpreter = PDFPageInterpreter(self.resource_manager, self.device)
            # pdf的page对象列表
            self.doc_pdfs = list(self.doc.get_pages())
        #  打开PDF文件, 生成一个包含图片doc对象的可迭代对象
        self.doc_pics = fitz.open(filename)
        self.pic_info = {}

    def to_pic(self, doc, zoom, pg, pic_path):
        """
        将单页pdf转换为pic
        :param doc: 图片的doc对象
        :param zoom: 图片缩放比例, type int, 数值越大分辨率越高
        :param pg: 对象在doc_pics中的索引
        :param pic_path: 图片保存路径
        :return: 图片的路径
        """
        rotate = int(0)
        trans = fitz.Matrix(zoom, zoom).preRotate(rotate)
        pm = doc.getPixmap(matrix=trans, alpha=False)
        path = os.path.join(pic_path, str(pg)) + '.png'
        pm.writePNG(path)
        return path
    
    def get_level(self, zip_loc_list):
        """
        记录每页PDF上图表的相对水平位置
        从上到下依次为level_0，level_1，...，level_n
        可以有几个图表在同一水平面上，即level_k(k=0,1,...,n)相同
        :param zip_loc_list: 包含loc_top和loc_bottom，这两个数组的元素一一对应
        示例：
        [(((33.8, 727.6025, 152.91049999999998, 737.4499999999999), '表 9：微盟 – 收入利润表'), 
        ((33.250054999999996, 388.634, 198.79753000000002, 396.95000000000005), '资料来源：公司财报，安信证券研究中心预测')), 
        (((33.8, 312.48249999999996, 158.16049999999998, 322.33), '表 10：微盟 – 现金流量表'), 
        ((33.250054999999996, 123.634, 198.79753000000002, 131.95), '资料来源：公司财报，安信证券研究中心预测'))]
        :return: 返回字典，key为相对水平位置，value为对应的图表列表，列表中的值是zip_loc_list中对应图表的index
        示例：
        {'level_0': [0], 'level_1': [1], 'level_2': [2, 3]}
        """
        level_dict = {}
        length = len(zip_loc_list)
        visit = [0] * length
        level_count = 0
        for i in range(length):
            # 为1表示对应的图表已经确定好其相对水平位置了，直接跳过
            if visit[i] == 1:
                continue
            # i对应的图表处在与之前确定位置的图表不同的水平位置上
            visit[i] = 1
            level_count_str = 'level_' + str(level_count)
            level_count += 1
            level_dict[level_count_str] = [i]
            # 排在i之前的对应的图表都已经确定好其相对水平位置了，不用再遍历了
            for j in range(i + 1, length):
                # 为1表示对应的图表已经确定好其相对水平位置了，直接跳过
                if visit[j] == 1:
                    continue

                # 将两个待比较的图表中更扁平的图表对应的高度值保存下来
                high = min(abs(zip_loc_list[i][0][0][3] - zip_loc_list[i][1][0][1]), \
                    abs(zip_loc_list[j][0][0][3] - zip_loc_list[j][1][0][1]))
                
                # 如果一个图表把另一个图表完全包住，即上边界比另一个高，下边界比另一个低，那么认为两个图表在一条水平线上
                if zip_loc_list[i][0][0][3] <= zip_loc_list[j][0][0][3] and \
                    zip_loc_list[i][1][0][1] >= zip_loc_list[j][1][0][1]:
                    visit[j] = 1
                    level_dict[level_count_str].append(j)
                
                elif zip_loc_list[i][0][0][3] >= zip_loc_list[j][0][0][3] and \
                    zip_loc_list[i][1][0][1] <= zip_loc_list[j][1][0][1]:
                    visit[j] = 1
                    level_dict[level_count_str].append(j)
                
                # 如果两个图表上边界高度之差和下边界高度之差都大于high值，那么认为两个图表不在一条水平线上
                elif min(abs(zip_loc_list[i][0][0][3] - zip_loc_list[j][0][0][3]), \
                    abs(zip_loc_list[i][1][0][1] - zip_loc_list[j][1][0][1])) >= high:
                    continue
                
                # 其他情况则认为两个图表在一条水平线上
                else:
                    visit[j] = 1
                    level_dict[level_count_str].append(j)
        return level_dict

    def get_pic_loc(self, doc, pgn):
        """
        获取单页中图片的位置
        :param doc: pdf的doc对象
        :return: 返回一个list, 元素为图片名称和上下y坐标元组组成的tuple. 当前页的尺寸
        """
        self.interpreter.process_page(doc)
        layout = self.device.get_result()
        # pdf的尺寸, tuple, (width, height)
        canvas_size = layout.bbox
        # 图片名称坐标
        loc_top = []
        # 来源坐标
        loc_bottom = []
        # left和right用于记录单页所有LT对象的最左和最右坐标
        left = canvas_size[2]
        right = 0
        # text_order用于记录当前页PDF中第一个含有文本的LT对象
        # first_text用于记录其文本内容
        text_order = 0
        first_text = ''
        # 遍历单页的所有LT对象
        for i in layout:
            left = min(left, i.bbox[0])
            right = max(right, i.bbox[2])
            if hasattr(i, 'get_text'):
                text = i.get_text().strip()
                text_order += 1
                if text_order == 1:
                    first_text = text
                # 匹配关键词
                if re.search(r'[图表]+\s*\d+[:：\s]*', text):
                    title_start = re.search(r'[图表]+\s*\d+[:：\s]*', text).start()
                    loc_top.append((i.bbox, text[title_start: ].replace('\n', '')))
                elif re.search(r'来源[:：\s]', text):
                    text = text.split('\n')[0]
                    loc_bottom.append((i.bbox, text))
                    # 如果单页得到的第一个关键词是r'来源[:：\s]'而不是r'[图表]+\s*\d+[:：\s]*'
                    # 那么有可能这是一个跨页的图表，需要去上一页中找是否有对应的关键词r'[图表]+\s*\d+[:：\s]*'
                    if len(loc_bottom) > 0 and len(loc_top) == 0:
                        last_pgn = pgn - 1
                        # self.pic_info[last_pgn]['loc_bottom']) < len(self.pic_info[last_pgn]['loc_top']
                        # 说明上一页存在关键词r'[图表]+\s*\d+[:：\s]*'没有与对应的关键词r'来源[:：\s]'配对
                        # 可以认为上一页关键词r'[图表]+\s*\d+[:：\s]*'以下到底边部分对应的是跨页图表的上半部分
                        # 本页关键词r'来源[:：\s]'以上到顶边部分对应的是跨页图表的下半部分
                        if last_pgn in self.pic_info and \
                            len(self.pic_info[last_pgn]['loc_bottom']) < len(self.pic_info[last_pgn]['loc_top']):
                            # 将上一页底边的坐标保存下来
                            self.pic_info[last_pgn]['loc_bottom'].append(((0, 0, 0, 0), ''))
                            # 将本页顶边的坐标保存下来
                            loc_top.append(((0, canvas_size[3], 0, canvas_size[3]), \
                                self.pic_info[last_pgn]['loc_top'][-1][1] + '@~@continue'))
                        # 如果上一页不存在关键词r'[图表]+\s*\d+[:：\s]*'与本页对应的关键词r'来源[:：\s]'配对
                        # 那么需要把关键词r'来源[:：\s]'去除掉
                        else:
                            if text_order == 1:  
                                loc_top.append(((0, canvas_size[3], 0, canvas_size[3]), first_text))
                            else:
                                loc_bottom.pop(-1)

        return (loc_top, loc_bottom, canvas_size, left, right)
        
    def get_crops(self, pic_path, canvas_size, position, cropped_pic_name, cropped_pic_path):
        """
        按给定位置截取图片
        :param pic_path: 被截取的图片的路径
        :param canvas_size: 图片为pdf时的尺寸, tuple, (0, 0, width, height)
        :param position: 要截取的位置, tuple, (x1, y1, x2, y2)
        :param cropped_pic_name: 截取的图片名称
        :param cropped_pic_path: 截取的图片保存路径
        :return:
        """
        img = Image.open(pic_path)
        # 当前图片的尺寸 tuple(width, height)
        pic_size = img.size
        # 截图的范围扩大值
        size_increase = 10
        x1 = max(pic_size[0] * (position[0] - size_increase)/canvas_size[2], 0)
        x2 = min(pic_size[0] * (position[2] + size_increase)/canvas_size[2], pic_size[0])
        y1 = max(pic_size[1] * (1 - (position[3] + size_increase)/canvas_size[3]), 0)
        y2 = min(pic_size[1] * (1 - (position[1] - size_increase)/canvas_size[3]), pic_size[1])
        try:
            cropped_img = img.crop((x1, y1, x2, y2))
            # 保存截图文件的路径
            path = os.path.join(cropped_pic_path, cropped_pic_name) + '.png'
            cropped_img.save(path)
            print('成功截取图片:', cropped_pic_name)
        except Exception as e:
            print(e)

    def get_pic_info(self, pic_path, page_count):
        """
        将PDF按页转成图片，并将相关信息按页保存在字典self.pic_info中
        key为PDF的页码，value为对应的相关信息
        :param pic_path: 被截取的图片路径
        :param page_count: PDF的页数
        :return:
        """
        if page_count <= 0:
            return
        for pgn in range(page_count):
            # 获取当前页的doc
            doc_pdf = self.doc_pdfs[pgn]
            doc_pic = self.doc_pics[pgn]
            # 将当前页转换为PNG, 返回值为图片路径
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
        对PDF按页提取图表。图表是从PDF转成的图片上截取下来的，不是直接在PDF进行操作
        :param cropped_pic_path: 截取的图片保存路径
        :return:
        """
        for k, v in self.pic_info.items():
            loc_list = list(zip(v['loc_top'], v['loc_bottom']))
            # loc_list为空表示该页PDF没有图表，直接跳过
            if not loc_list:
                continue
            width = abs(v['right'] - v['left'])
            # 将图表列表按照图表的x1坐标升序重排
            loc_list.sort(key=lambda x: x[0][0][0])
            path = v['path']
            canvas_size = v['canvas_size']
            level_dict = self.get_level(loc_list)
            loc_name_pic = []

            for level, item_list in level_dict.items():
                level_count = len(item_list)
                level_order = 0
                # 一个水平面上，有几个图表就将width几等分，用于计算x1和x2坐标
                for item in item_list:
                    x1 = v['left'] + level_order * width / level_count
                    if level_order > 0:
                        x1 = min(min(x1, loc_list[item_list[level_order]][0][0][0]), \
                            loc_list[item_list[level_order]][0][0][2])
                    level_order += 1
                    x2 = v['left'] + level_order * width / level_count
                    # 如果一个水平面上有多个图表，并且当前图表不是最右边的一个图表
                    # 令x2与当前图表右边图表的x1坐标进行比较，取较小值做为当前图表的x2坐标
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
        对得到的图表进行遍历，将跨页的图表的上下两部分拼接成一张图表保存，并删除原来的上下两部分图表
        :param cropped_pic_path: 截取的图片保存路径
        :return:
        """
        pic_count_dict = {}
        pic_list = os.listdir(cropped_pic_path)
        # 将图片按照名字保存到字典pic_count_dict中，如果是跨页的图表，同一个名字会对应多张图片
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
            # 避免上下顺序出错，加了一个判断
            if '@~@continue' in png1:
                png1, png2 = png2, png1
            img1, img2 = Image.open(png1), Image.open(png2)
            size1, size2 = img1.size, img2.size
            # 新图片的宽为两张图的较小值，高为两张图的高之和
            joint = Image.new('RGB', (min(size1[0], size2[0]), size1[1] + size2[1]))
            loc1, loc2 = (0, 0), (0, size1[1])
            joint.paste(img1, loc1)
            joint.paste(img2, loc2)
            # 删除原来的图片，保存新图片
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