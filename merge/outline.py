# -*- coding: utf-8 -*-
"""
HWPX 개요 트리 관련 함수

개요:
- build_outline_tree: 문단 리스트를 개요 트리로 변환
- merge_outline_trees: 여러 개요 트리 병합
- filter_outline_tree: 특정 개요 제외
- flatten_outline_tree: 개요 트리를 문단 리스트로 펼침
"""

from typing import List, Dict, Optional, Set, Tuple
import copy

from .models import Paragraph, OutlineNode


def build_outline_tree(paragraphs: List[Paragraph]) -> List[OutlineNode]:
    """문단 리스트를 개요 트리로 변환"""
    if not paragraphs:
        return []

    root_nodes: List[OutlineNode] = []
    node_stack: List[Tuple[int, OutlineNode]] = []
    pre_outline_paras: List[Paragraph] = []

    for para in paragraphs:
        if para.is_outline:
            node = OutlineNode(
                level=para.level,
                name=para.text,
                paragraphs=[para],
            )

            if pre_outline_paras and not root_nodes and not node_stack:
                node.paragraphs = pre_outline_paras + node.paragraphs
                pre_outline_paras = []

            while node_stack and node_stack[-1][0] >= para.level:
                node_stack.pop()

            if node_stack:
                parent = node_stack[-1][1]
                parent.children.append(node)
            else:
                root_nodes.append(node)

            node_stack.append((para.level, node))
        else:
            if node_stack:
                current_node = node_stack[-1][1]
                current_node.paragraphs.append(para)
            else:
                pre_outline_paras.append(para)

    if not root_nodes and pre_outline_paras:
        node = OutlineNode(
            level=-1,
            name="",
            paragraphs=pre_outline_paras,
        )
        root_nodes.append(node)

    return root_nodes


def merge_outline_trees(
    trees_list: List[List[OutlineNode]],
    exclude_outlines: Optional[Set[str]] = None
) -> List[OutlineNode]:
    """
    여러 개요 트리를 병합

    Args:
        trees_list: 병합할 개요 트리 리스트
        exclude_outlines: 제외할 개요 이름 집합 (예: {"1. 서론", "3.2 실험 방법"})
                          정확한 이름 매칭 또는 접두사 매칭 지원
                          예: "3." → "3."으로 시작하는 모든 개요 제외

    Returns:
        병합된 개요 트리
    """
    if not trees_list:
        return []

    if len(trees_list) == 1:
        result = copy.deepcopy(trees_list[0])
        if exclude_outlines:
            result = filter_outline_tree(result, exclude_outlines)
        return result

    merged = copy.deepcopy(trees_list[0])

    for tree in trees_list[1:]:
        merged = _merge_two_trees(merged, tree, exclude_outlines)

    # 최종 결과에서도 제외 적용
    if exclude_outlines:
        merged = filter_outline_tree(merged, exclude_outlines)

    return merged


def _merge_two_trees(
    base: List[OutlineNode],
    other: List[OutlineNode],
    exclude_outlines: Optional[Set[str]] = None
) -> List[OutlineNode]:
    """두 개요 트리 병합"""
    result = copy.deepcopy(base)

    name_to_idx: Dict[str, int] = {}
    for i, node in enumerate(result):
        if node.name:
            name_to_idx[node.name] = i

    for other_node in other:
        # 제외 대상인지 확인
        if exclude_outlines and _should_exclude(other_node.name, exclude_outlines):
            continue

        other_node = copy.deepcopy(other_node)

        if other_node.name in name_to_idx:
            idx = name_to_idx[other_node.name]
            base_node = result[idx]

            # Template의 첫 번째 내용 문단 스타일 가져오기
            base_content = base_node.get_content_paragraphs()
            template_style = base_content[0].para_pr_id if base_content else None

            content_paras = other_node.get_content_paragraphs()

            # Addition 문단에 Template 스타일 적용
            if template_style:
                for para in content_paras:
                    para.para_pr_id = template_style

            base_node.paragraphs.extend(content_paras)

            base_node.children = _merge_two_trees(base_node.children, other_node.children, exclude_outlines)
        else:
            result.append(other_node)
            if other_node.name:
                name_to_idx[other_node.name] = len(result) - 1

    return result


def _should_exclude(name: str, exclude_outlines: Set[str]) -> bool:
    """개요 이름이 제외 대상인지 확인"""
    if not name or not exclude_outlines:
        return False

    for pattern in exclude_outlines:
        # 정확한 매칭
        if name == pattern:
            return True
        # 접두사 매칭 (패턴이 "."으로 끝나면 접두사 매칭)
        if pattern.endswith('.') and name.startswith(pattern):
            return True
        # 패턴이 숫자로 시작하고 개요도 같은 숫자로 시작하면 매칭
        # 예: "3" → "3. 실험", "3.1 방법" 등 매칭
        if pattern.isdigit() and name.startswith(pattern + '.'):
            return True

    return False


def filter_outline_tree(
    nodes: List[OutlineNode],
    exclude_outlines: Set[str]
) -> List[OutlineNode]:
    """
    개요 트리에서 특정 개요 제외

    Args:
        nodes: 개요 노드 리스트
        exclude_outlines: 제외할 개요 이름 집합

    Returns:
        필터링된 개요 트리
    """
    if not exclude_outlines:
        return nodes

    result = []
    for node in nodes:
        # 제외 대상인지 확인
        if _should_exclude(node.name, exclude_outlines):
            continue

        # 노드 복사
        filtered_node = OutlineNode(
            level=node.level,
            name=node.name,
            paragraphs=list(node.paragraphs),
            children=filter_outline_tree(node.children, exclude_outlines)
        )
        result.append(filtered_node)

    return result


def get_all_outline_names(nodes: List[OutlineNode], indent: str = "") -> List[Tuple[str, str]]:
    """
    개요 트리에서 모든 개요 이름 추출 (선택 UI용)

    Args:
        nodes: 개요 노드 리스트
        indent: 들여쓰기 문자열

    Returns:
        (표시용 문자열, 실제 이름) 튜플 리스트
    """
    result = []
    for node in nodes:
        if node.name:
            display = f"{indent}[{node.level}] {node.name}"
            result.append((display, node.name))

        if node.children:
            child_names = get_all_outline_names(node.children, indent + "  ")
            result.extend(child_names)

    return result


def flatten_outline_tree(nodes: List[OutlineNode]) -> List[Paragraph]:
    """개요 트리를 문단 리스트로 펼침"""
    result = []
    for node in nodes:
        result.extend(node.paragraphs)
        result.extend(flatten_outline_tree(node.children))
    return result


def print_outline_tree(nodes: List[OutlineNode], indent: int = 0):
    """개요 트리 출력"""
    indent_str = "  " * indent

    for node in nodes:
        if node.level >= 0:
            print(f"{indent_str}[level {node.level}] {node.name}")

            content_count = len(node.get_content_paragraphs())
            if content_count > 0:
                print(f"{indent_str}  - 내용 문단: {content_count}개")
        else:
            print(f"{indent_str}[일반 문단 {len(node.paragraphs)}개]")

        if node.children:
            print_outline_tree(node.children, indent + 1)
