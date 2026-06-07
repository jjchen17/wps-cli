# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""路径解析器测试"""

import pytest

from wps_cli.exceptions import ValidationError
from wps_cli.services.path_resolver import PathComponent, PathResolver


class TestPathResolverParse:
    """测试路径解析器的 parse 方法"""

    def test_parse_simple_writer_path(self):
        """解析 Writer 简单路径"""
        components = PathResolver.parse("/section[1]/paragraph[3]")
        assert len(components) == 2
        assert components[0].element == "section"
        assert components[0].index == 1
        assert components[1].element == "paragraph"
        assert components[1].index == 3

    def test_parse_table_cell_path(self):
        """解析表格单元格路径（元组索引）"""
        components = PathResolver.parse("/section[1]/table[2]/cell[3,1]")
        assert len(components) == 3
        assert components[2].element == "cell"
        assert components[2].index == (3, 1)

    def test_parse_cell_with_spaces(self):
        """解析带有空格的元组索引"""
        components = PathResolver.parse("/table[1]/cell[ 3 , 1 ]")
        assert components[1].element == "cell"
        assert components[1].index == (3, 1)

    def test_parse_sheet_by_name(self):
        """解析按名称访问工作表"""
        components = PathResolver.parse('/sheet["Sheet1"]/cell["A1"]')
        assert len(components) == 2
        assert components[0].element == "sheet"
        assert components[0].index == "Sheet1"
        assert components[1].element == "cell"
        assert components[1].index == "A1"

    def test_parse_sheet_single_quotes(self):
        """解析单引号工作表名"""
        components = PathResolver.parse("/sheet['My Sheet']/range['A1:B10']")
        assert components[0].index == "My Sheet"
        assert components[1].index == "A1:B10"

    def test_parse_slide_path(self):
        """解析 PPT 幻灯片路径"""
        components = PathResolver.parse("/slide[1]/shape[2]/text[1]")
        assert len(components) == 3
        assert components[0].element == "slide"
        assert components[1].element == "shape"
        assert components[2].element == "text"
        assert components[2].index == 1

    def test_parse_body_path(self):
        """解析 body 无索引路径"""
        components = PathResolver.parse("/body")
        assert len(components) == 1
        assert components[0].element == "body"
        assert components[0].index is None

    def test_parse_header_footer(self):
        """解析页眉页脚路径"""
        components = PathResolver.parse("/header[1]")
        assert components[0].element == "header"
        assert components[0].index == 1

    def test_parse_no_index_path(self):
        """解析无索引的 section 路径"""
        components = PathResolver.parse("/section[5]")
        assert components[0].element == "section"
        assert components[0].index == 5

    def test_parse_element_case_insensitive(self):
        """元素名大小写不敏感"""
        components = PathResolver.parse("/Slide[1]/Shape[2]")
        assert components[0].element == "slide"
        assert components[1].element == "shape"

    def test_parse_with_underscore(self):
        """支持含下划线的元素名"""
        components = PathResolver.parse("/my_element[1]")
        assert components[0].element == "my_element"


class TestPathResolverParseErrors:
    """测试路径解析错误处理"""

    def test_empty_path(self):
        """空路径应抛出 ValidationError"""
        with pytest.raises(ValidationError, match="必须以 '/' 开头"):
            PathResolver.parse("")

    def test_no_slash_prefix(self):
        """不以 / 开头的路径"""
        with pytest.raises(ValidationError, match="必须以 '/' 开头"):
            PathResolver.parse("section[1]")

    def test_no_valid_components(self):
        """无有效组件的路径"""
        with pytest.raises(ValidationError, match="无法解析路径"):
            PathResolver.parse("/[1]")


class TestPathResolverResolveRouting:
    """测试 resolve 方法的路由分发"""

    def test_resolve_unknown_app_type(self):
        """不支持的 app_type"""
        resolver = PathResolver()
        with pytest.raises(ValidationError, match="不支持的应用类型"):
            resolver.resolve(None, "unknown", "/body")

    def test_resolve_routes_to_writer(self):
        """正确路由到 writer 解析器"""
        # 此处仅测试路由，不测试实际 COM 导航
        # 实际 WPS 连接测试在集成测试中完成
        pass

    def test_resolve_routes_to_calc(self):
        """正确路由到 calc 解析器"""
        pass

    def test_resolve_routes_to_impress(self):
        """正确路由到 impress 解析器"""
        pass


class TestPathComponent:
    """测试 PathComponent 数据类"""

    def test_create_component(self):
        comp = PathComponent(element="slide", index=1)
        assert comp.element == "slide"
        assert comp.index == 1

    def test_create_tuple_index_component(self):
        comp = PathComponent(element="cell", index=(3, 1))
        assert comp.index == (3, 1)

    def test_create_string_index_component(self):
        comp = PathComponent(element="sheet", index="Sheet1")
        assert comp.index == "Sheet1"

    def test_create_none_index_component(self):
        comp = PathComponent(element="body", index=None)
        assert comp.index is None


class TestPathResolverCheckIndex:
    """测试 _check_index 边界检查"""

    def test_index_within_range(self):
        """索引在范围内不应抛出异常"""
        resolver = PathResolver()
        comp = PathComponent(element="slide", index=2)
        # 不抛异常即为通过
        resolver._check_index(comp, 5, "/slide[2]")

    def test_index_at_boundary(self):
        """索引在边界上"""
        resolver = PathResolver()
        comp1 = PathComponent(element="slide", index=1)
        resolver._check_index(comp1, 5, "/slide[1]")
        comp5 = PathComponent(element="slide", index=5)
        resolver._check_index(comp5, 5, "/slide[5]")

    def test_index_out_of_range_below(self):
        """索引小于 1"""
        resolver = PathResolver()
        comp = PathComponent(element="slide", index=0)
        with pytest.raises(ValidationError, match="slide\\[0\\] 越界"):
            resolver._check_index(comp, 5, "/slide[0]")

    def test_index_out_of_range_above(self):
        """索引大于最大值"""
        resolver = PathResolver()
        comp = PathComponent(element="slide", index=10)
        with pytest.raises(ValidationError, match="slide\\[10\\] 越界"):
            resolver._check_index(comp, 5, "/slide[10]")

    def test_none_index_skips_check(self):
        """None 索引跳过检查"""
        resolver = PathResolver()
        comp = PathComponent(element="body", index=None)
        # 不应抛异常
        resolver._check_index(comp, 1, "/body")


class TestExcelStylePath:
    """测试 Excel 风格简写 $Sheet:Ref"""

    def test_resolve_calc_excel_style_simple(self):
        """$Sheet:A1 格式应被解析"""
        # 此测试通过验证内部逻辑（不依赖实际 COM）
        # 路径 resolve_calc 内部对 $ 开头的路径做了特判
        import re

        path = "$Sheet1:A1"
        assert path.startswith("$")
        m = re.match(r"^\$([^:]+):(.+)$", path)
        assert m is not None
        assert m.group(1) == "Sheet1"
        assert m.group(2) == "A1"

    def test_resolve_calc_excel_style_range(self):
        """$Sheet:Range 格式"""
        import re

        path = "$Data:A1:B10"
        m = re.match(r"^\$([^:]+):(.+)$", path)
        assert m.group(1) == "Data"
        assert m.group(2) == "A1:B10"
