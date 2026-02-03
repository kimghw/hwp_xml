# -*- coding: utf-8 -*-
"""
Claude Agent SDK 형식 검토 모듈

HWPX 문서의 형식(캡션, 글머리 기호)을 Agent를 통해 검토합니다.

사용 예:
    from agent.format_review import merge_with_review_async, FORMAT_REVIEW_PROMPT

    result = await merge_with_review_async(
        hwpx_paths=["file1.hwpx", "file2.hwpx"],
        output_path="output.hwpx",
        use_agent=True
    )
"""

from pathlib import Path
from typing import List, Union


FORMAT_REVIEW_PROMPT = """당신은 HWPX 문서 형식 검토 전문가입니다.

다음 규칙에 따라 문서 형식을 검토하세요:

## 캡션 스타일 규칙
- 테이블: "표 N. 제목" 형식 (예: "표 1. 실험 결과")
- 그림: "그림 N. 제목" 형식 (예: "그림 1. 시스템 구조도")

## 개요 글머리 기호 규칙 (3단계)
- 1단계: ■ (검정 네모)
- 2단계: ● (검정 원)
- 3단계: - (대시)

## 검토 시 확인할 사항
1. 모든 캡션이 표준 형식을 따르는지
2. 글머리 기호가 계층 구조에 맞는지
3. 번호가 연속적인지

검토 결과를 JSON 형식으로 반환하세요:
{
    "is_valid": true/false,
    "errors": [...],
    "warnings": [...],
    "suggestions": [...]
}"""


async def merge_with_review_async(
    hwpx_paths: List[Union[str, Path]],
    output_path: Union[str, Path],
    use_agent: bool = True
):
    """
    Claude Agent SDK를 사용한 비동기 병합 + 검토

    Args:
        hwpx_paths: 병합할 HWPX 파일 경로들
        output_path: 출력 파일 경로
        use_agent: Claude Agent 사용 여부

    Returns:
        MergeResult
    """
    from merge.merge_pipeline import MergePipeline, MergeResult

    pipeline = MergePipeline()

    if not use_agent:
        return pipeline.merge(hwpx_paths, output_path, auto_fix=True)

    # SDK import 시도
    try:
        from claude_code_sdk import query, ClaudeCodeOptions
    except ImportError as e:
        print(f"경고: claude-code-sdk import 실패 ({e}). 동기 방식으로 진행합니다.")
        return pipeline.merge(hwpx_paths, output_path, auto_fix=True)

    result = MergeResult(output_path=str(output_path))

    try:
        # Agent 옵션 설정
        options = ClaudeCodeOptions(
            allowed_tools=["Read", "Grep", "Glob"],
            cwd=str(Path(hwpx_paths[0]).parent),
        )

        # 1단계: Agent로 형식 검토 요청
        print("[Agent] 형식 검토 시작...")

        file_list = '\n'.join(f'- {p}' for p in hwpx_paths)
        review_prompt = f"""다음 HWPX 파일들의 형식을 검토해주세요:
{file_list}

{FORMAT_REVIEW_PROMPT}
"""

        agent_responses = []
        async for message in query(prompt=review_prompt, options=options):
            if hasattr(message, 'content'):
                for block in message.content:
                    if hasattr(block, 'text'):
                        agent_responses.append(block.text)
            elif hasattr(message, 'result'):
                agent_responses.append(str(message.result))

        result.agent_feedback = '\n'.join(agent_responses)
        if result.agent_feedback:
            print(f"[Agent 검토 결과]\n{result.agent_feedback[:500]}...")

    except Exception as e:
        print(f"[Agent 오류] {e}")
        result.agent_feedback = f"Agent 실행 중 오류: {e}"

    # 2단계: 로컬 병합 + 수정
    print("\n[로컬] 병합 시작...")
    merge_result = pipeline.merge(hwpx_paths, output_path, auto_fix=True)

    # 결과 병합
    result.validation = merge_result.validation
    result.success = merge_result.success
    result.fixes_applied = merge_result.fixes_applied

    return result
