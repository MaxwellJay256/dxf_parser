import ezdxf
from vSDK_ShapeTools import ShapeEditor
from vSDK import *
import logging
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

def vSDK_Layer_ExportGerber(Pcb, LayerID: int, CExportGerberFileName, Mode: int):
    vSDK_dll.vSDK_Layer_ExportGerber(Pcb, LayerID, CExportGerberFileName, Mode)

layer_name_default = "CustomLayer"
job_path_default = b"E:/DFXMetaLabDev/jobs/CustomLayer.vayo/CustomLayer.job"
dxf_path_default = "E:/DFXMetaLabDev/scripts/dxf_parser/Top.dxf"

class DXFParser():
    def __init__(self, sdk_path):
        self.sdk_path = sdk_path
        self.shape_editor = None

        self.root = tk.Tk()
        self.setup_ui(self.root)
        self.root.mainloop()

    def setup_ui(self, root):
        root.title("DXF Parser")
        root.geometry("500x400")
        root.minsize(400, 280)
        root.resizable(True, True)

        default_font = ("微软雅黑", 10)
        root.option_add("*Font", default_font)

        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        label_title = ttk.Label(main_frame, text="DXF Parser", anchor="center", font=("微软雅黑", 12))
        label_title.pack(fill=tk.X, pady=(0, 10))

        # Job 文件选择
        frame_job_path = ttk.Frame(main_frame)
        frame_job_path.pack(fill=tk.X, pady=5)
        label_job_path = ttk.Label(frame_job_path, text="Job 文件：", width=9)
        label_job_path.pack(side=tk.LEFT)
        self.entry_job_path = ttk.Entry(frame_job_path, width=30)
        self.entry_job_path.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.entry_job_path.insert(0, job_path_default.decode('utf-8'))  # 填充默认 Job 路径
        button_job_path = ttk.Button(frame_job_path, text="...", width=3, command=self.select_job_path)
        button_job_path.pack(side=tk.RIGHT)

        # DXF 文件选择
        frame_dxf_path = ttk.Frame(main_frame)
        frame_dxf_path.pack(fill=tk.X, pady=5)
        label_dxf_path = ttk.Label(frame_dxf_path, text="DXF 文件：", width=9)
        label_dxf_path.pack(side=tk.LEFT)
        self.entry_dxf_path = ttk.Entry(frame_dxf_path, width=30)
        self.entry_dxf_path.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.entry_dxf_path.insert(0, dxf_path_default)  # 填充默认 DXF 路径
        button_dxf_path = ttk.Button(frame_dxf_path, text="...", width=3, command=self.select_dxf_path)
        button_dxf_path.pack(side=tk.RIGHT)

        # 层名输入框
        frame_layer_name = ttk.Frame(main_frame)
        frame_layer_name.pack(fill=tk.X, pady=5)
        label_layer_name = ttk.Label(frame_layer_name, text="目标层名：", width=9)
        label_layer_name.pack(side=tk.LEFT)
        self.entry_layer_name = ttk.Entry(frame_layer_name, width=30)
        self.entry_layer_name.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.entry_layer_name.insert(0, layer_name_default)  # 填充默认层名

        # 垂直填充
        frame_spacer = ttk.Frame(main_frame)
        frame_spacer.pack(fill=tk.BOTH, expand=True)

        # 按钮
        frame_buttons = ttk.Frame(main_frame)
        frame_buttons.pack(fill=tk.X, pady=10)
        button_parse = ttk.Button(frame_buttons, text="解析 DXF", command=self.parse_dxf)
        button_parse.pack(padx=5, fill=tk.X, expand=True)

    def select_job_path(self):
        '''
        弹出文件选择对话框，选择 Job 文件
        '''
        path = filedialog.askopenfilename(title="选择 Job 文件", filetypes=[("Job 文件", "*.job"), ("所有文件", "*.*")])
        if path:
            self.entry_job_path.delete(0, tk.END)
            self.entry_job_path.insert(0, path)

    def select_dxf_path(self):
        '''
        弹出文件选择对话框，选择 DXF 文件
        '''
        path = filedialog.askopenfilename(title="选择 DXF 文件", filetypes=[("DXF 文件", "*.dxf"), ("所有文件", "*.*")])
        if path:
            self.entry_dxf_path.delete(0, tk.END)
            self.entry_dxf_path.insert(0, path)

    def load_dxf(self, dxf_path: str):
        '''
        加载 DXF 文件，并将模型空间中的实体存储到 msp 属性中。
        
        :param dxf_path: DXF 文件的绝对路径
        '''
        if not os.path.exists(dxf_path):
            raise FileNotFoundError(f"DXF file does not exist: {dxf_path}")
        doc = ezdxf.readfile(dxf_path)
        logging.info(f"DXF file loaded: {dxf_path}")
        self.msp = doc.modelspace()

    def get_dxf_layers(self) -> list:
        '''
        获取 DXF 文件中的所有层名称
        
        :return: 包含所有层名的列表
        '''
        layers = set()
        for entity in self.msp:
            layer_name = entity.dxf.layer
            layers.add(layer_name)
        return list(layers)

    def parse_circle(self, entity, target_layer_id: int) -> int:
        '''
        解析 CIRCLE 实体，并绘制为空心圆。
        
        :param entity: CIRCLE 实体
        :param target_layer_id: 目标层 ID
        :return: 绘制的圆形对象 ID
        '''
        if not entity.dxftype() == 'CIRCLE':
            error_msg = f"Entity is not a CIRCLE: {entity.dxftype()}"
            raise ValueError(error_msg)
        
        center = entity.dxf.center
        radius = entity.dxf.radius
        logging.info(f"Drawing CIRCLE at {center} with radius {radius}")
        diameter = radius * 2
        return self.shape_editor.circle(center[0], center[1], diameter, target_layer_id, circleFilled=False)
    
    def parse_line(self, entity, target_layer_id: int) -> int:
        '''
        解析 LINE 实体，并绘制为线段。
        
        :param entity: LINE 实体
        :param target_layer_id: 目标层 ID
        :return: 绘制的线段对象 ID
        '''
        if not entity.dxftype() == 'LINE':
            error_msg = f"Entity is not a LINE: {entity.dxftype()}"
            raise ValueError(error_msg)
        
        start = entity.dxf.start
        end = entity.dxf.end
        logging.info(f"Drawing line from {start} to {end}")
        
        return self.shape_editor.line(start[0], start[1], end[0], end[1], target_layer_id)

    def parse_lwpolyline(self, entity, target_layer_id: int) -> int:
        '''
        解析 LWPOLYLINE 实体，并根据其点的特征绘制为多段线段、闭合多边形或弧线。
        
        :param entity: LWPOLYLINE 实体
        :param target_layer_id: 目标层 ID
        :return: 绘制的对象 ID
        '''
        if not entity.dxftype() == 'LWPOLYLINE':
            erroe_msg = f"Entity is not a LWPOLYLINE: {entity.dxftype()}"
            raise ValueError(erroe_msg)
        
        points = entity.get_points()
        if len(points) > 3 and points[0][:2] == points[-1][:2]:
            # 如果第 1 个点与最后一个点相同，则认为是闭合多边形，交给 parse_lwpolyline_closed_polygon 处理
            return self.parse_lwpolyline_closed_polygon(entity, target_layer_id)
        elif len(points) == 2 and all(p[4] == 1.0 for p in entity):
            # 如果只含 2 个点，且 2 个点的 bulge 值均为 1.0，则认为是弧线，交给 parse_lwpolyline_arc 处理
            return self.parse_lwpolyline_arc(entity, target_layer_id)
        else:
            # 剩余的情况认为是多段线段，交给 parse_lwpolyline_line 处理
            return self.parse_lwpolyline_line(entity, target_layer_id)

    def parse_lwpolyline_closed_polygon(self, entity, target_layer_id: int) -> int:
        '''
        解析 LWPOLYLINE 实体，并绘制为闭合多边形。

        :param entity: LWPOLYLINE 实体
        :param target_layer_id: 目标层 ID
        :return: 绘制的多边形对象 ID
        '''
        if not entity.dxftype() == 'LWPOLYLINE':
            error_msg = f"Entity is not a LWPOLYLINE: {entity.dxftype()}"
            raise ValueError(error_msg)
        
        points = entity.get_points()
        if len(points) <= 3:
            error_msg = "LWPOLYLINE must have at least 4 points for closed polygon parsing."
            raise ValueError(error_msg)
        
        logging.info(f"Drawing closed LWPOLYLINE polygon with points: {points}")
        # 忽略 points 中的最后一个点，绘制多边形
        return self.shape_editor.polygon(
            [point[:2] for point in points[:-1]],
            target_layer_id,
            Filled=False
        )

    def parse_lwpolyline_arc(self, entity, target_layer_id: int) -> int:
        '''
        解析 LWPOLYLINE 实体，并绘制为弧线。

        :param entity: LWPOLYLINE 实体
        :param target_layer_id: 目标层 ID
        :return: 绘制的弧线对象 ID
        '''
        if not entity.dxftype() == 'LWPOLYLINE':
            error_msg = f"Entity is not a LWPOLYLINE: {entity.dxftype()}"
            raise ValueError(error_msg)
        
        points = entity.get_points()
        if len(points) < 2:
            error_msg = "LWPOLYLINE must have at least 2 points for arc parsing."
            raise ValueError(error_msg)

        start_point = points[0][:2]  # 弧线起点
        end_point = points[-1][:2]  # 弧线终点
        bulge = points[0][-1]
        # 用 ezdxf 的 bulge_to_arc 方法计算圆心、起始角度、终止角度和半径
        center, start_angle, end_angle, radius = ezdxf.math.bulge_to_arc(start_point, end_point, bulge)
        angle_rotate = end_angle - start_angle
        logging.info(
            f"Drawing LWPOLYLINE arc with center: {center}, start angle: {start_angle}, end angle: {end_angle}, radius: {radius}"
        )
        return self.shape_editor.arc(center[0], center[1], radius, start_angle, angle_rotate, target_layer_id, Filled=False)

    def parse_lwpolyline_line(self, entity, target_layer_id: int) -> int:
        '''
        解析 LWPOLYLINE 实体，并绘制为多段线段。
        
        :param entity: LWPOLYLINE 实体
        :param target_layer_id: 目标层 ID
        :return: 绘制的最后一段线段对象 ID
        '''
        if not entity.dxftype() == 'LWPOLYLINE':
            error_msg = f"Entity is not a LWPOLYLINE: {entity.dxftype()}"
            raise ValueError(error_msg)
        
        points = entity.get_points()
        if len(points) < 2:
            error_msg = "LWPOLYLINE must have at least 2 points for line parsing."
            raise ValueError(error_msg)

        logging.info(f"Drawing LWPOLYLINE line with points: {points}")
        for i in range(len(points) - 1):
            start = points[i][:2]
            end = points[i + 1][:2]
            object_id = self.shape_editor.line(start[0], start[1], end[0], end[1], target_layer_id)
        return object_id
    
    def parse_hatch(self, entity, target_layer_id: int) -> int:
        '''
        解析 HATCH 实体，并绘制为填充的多边形。

        :param entity: HATCH 实体
        :param target_layer_id: 目标层 ID
        :return: 绘制的多边形对象 ID
        '''
        if not entity.dxftype() == 'HATCH':
            error_msg = f"Entity is not a HATCH: {entity.dxftype()}"
            raise ValueError(error_msg)
        
        # 统计 HATCH 边界的顶点
        vertices = []
        for path in entity.paths:
            for edge in path.edges:
                vertices.append(edge.start)  # 保存每个 edge 的起点作为多边形顶点

        if not vertices:
            warn_msg = "HATCH has no vertices to draw."
            raise ValueError(warn_msg)
        
        logging.info(f"Drawing HATCH with vertices: {vertices}")
        # 使用 polygon 方法绘制 HATCH 的边界
        return self.shape_editor.polygon(
            [(vertex.x, vertex.y) for vertex in vertices], target_layer_id, Filled=True
        )
    
    def parse_dxf(self):
        '''
        解析 DXF 文件中的所有实体，将 CIRCLE、LINE、LWPOLYLINE 和 HATCH 实体绘制到单个指定层。
        对于其余的实体类型，记录日志并跳过。
        '''
        job_path = self.entry_job_path.get()
        dxf_path = self.entry_dxf_path.get()
        if not job_path or not dxf_path:
            logging.warning("Job file or DXF file not selected.")
            messagebox.showwarning("Warning", "请先选择 Job 文件和 DXF 文件")
            return
        layer_name = self.entry_layer_name.get()
        if not layer_name:
            logging.warning("Layer name cannot be empty.")
            messagebox.showwarning("Warning", "层名称不能为空")
            return
        self.job_path = job_path.encode('utf-8')
        
        # 初始化 ShapeEditor
        self.shape_editor = ShapeEditor(self.sdk_path, self.job_path)

        try:
            self.load_dxf(dxf_path)
        except FileNotFoundError as e:
            logging.error(e)
            messagebox.showerror("Error", f"加载 DXF 文件失败: {e}")
            return
        
        # 在 job 中新建一个层
        target_layer_id = self.shape_editor.add_layer(layer_name, layer_side=True)[0]
        logging.info(f"Layer '{layer_name}' created with ID: {target_layer_id}")

        # 解析 DXF 中所有实体
        circle_count = 0
        line_count = 0
        lwpolyline_count = 0
        hatch_count = 0
        for entity in self.msp:
            if entity.dxftype() == 'CIRCLE':
                try:
                    self.parse_circle(entity, target_layer_id)
                    circle_count += 1
                except Exception as e:
                    logging.error(f"Failed to parse CIRCLE: {e}")
            elif entity.dxftype() == 'LINE':
                try:
                    self.parse_line(entity, target_layer_id)
                    line_count += 1
                except Exception as e:
                    logging.error(f"Failed to parse LINE: {e}")
            elif entity.dxftype() == 'LWPOLYLINE':
                try:
                    self.parse_lwpolyline(entity, target_layer_id)
                    lwpolyline_count += 1
                except Exception as e:
                    logging.error(f"Failed to parse LWPOLYLINE: {e}")
            elif entity.dxftype() == 'HATCH':
                try:
                    self.parse_hatch(entity, target_layer_id)
                    hatch_count += 1
                except Exception as e:
                    logging.error(f"Failed to parse HATCH: {e}")
            else:
                # 对于不支持的实体类型，记录日志并跳过
                logging.info(f"Unsupported DXF entity type: {entity.dxftype()}")
        vSDK_SaveJob()

        # 导出 Gerber
        job_folder = self.export_layer_gerber(layer_name, target_layer_id)
        if job_folder == "":
            messagebox.showerror("Error", "导出 Gerber 文件失败，请检查日志")
            return
    
        messagebox.showinfo(
            "Success",
            f"DXF 解析完成，导出对象统计\n"
            f"CIRCLE: {circle_count}\n"
            f"LINE: {line_count}\n"
            f"LWPOLYLINE: {lwpolyline_count}\n"
            f"HATCH: {hatch_count}\n"
            f"Gerber 文件已导出到：{job_folder}/{layer_name}.gbr"
        )

    def export_layer_gerber(self, layer_name: str, layer_id: int) -> str:
        '''
        导出指定层的 Gerber 文件到 Job 目录
        
        :param layer_name: 层名称
        :param layer_id: 层 ID
        :return: Gerber 文件路径
        '''
        layer = vSDK_Board_GetLayerByName(self.shape_editor.board, layer_name.encode())
        if not layer.value:
            logging.error("Layer %s not found", layer_name)
            return ""
        
        # 导出 Gerber 文件
        job_folder = os.path.dirname(self.job_path.decode('utf-8'))
        if not os.path.exists(job_folder):
            logging.error("Job folder %s does not exist.", job_folder)
            return ""
        gerber_path = os.path.join(job_folder, layer_name + ".gbr")  # 导出路径为 Job 目录
        vSDK_Layer_ExportGerber(self.shape_editor.pcb, layer_id, gerber_path.encode(), 0)
        logging.info("Exported Gerber files to %s", gerber_path)
        return job_folder
    
if __name__ == "__main__":
    sdk_path = b"C:/VayoPro/DFX_MetaLab"    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
    )
    parser = DXFParser(sdk_path)  # 实例化 DXFParser
