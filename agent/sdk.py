# -*- coding: utf-8 -*-
"""
Claude Code SDK 호출 모듈

Claude CLI를 통해 SDK 기능을 호출합니다.
"""

import subprocess
from typing import Optional
from dataclasses import dataclass


@dataclass
class SDKResult:
    """SDK 호출 결과"""
    success: bool = True
    output: str = ""
    error: str = ""


class ClaudeSDK:
    """Claude Code SDK 호출 클래스"""

    def __init__(self, timeout: int = 30):
        """
        Args:
            timeout: SDK 호출 타임아웃 (초)
        """
        self.timeout = timeout

    def call(self, prompt: str) -> SDKResult:
        """
        Claude SDK 호출

        Args:
            prompt: 프롬프트 텍스트

        Returns:
            SDKResult: 호출 결과
        """
        try:
            result = subprocess.run(
                ['claude', '-p', prompt],
                capture_output=True,
                text=True,
                timeout=self.timeout
            )

            if result.returncode == 0:
                return SDKResult(
                    success=True,
                    output=result.stdout.strip(),
                    error=""
                )
            else:
                return SDKResult(
                    success=False,
                    output="",
                    error=result.stderr.strip()
                )

        except subprocess.TimeoutExpired:
            return SDKResult(
                success=False,
                output="",
                error=f"SDK 호출 타임아웃 ({self.timeout}초)"
            )
        except FileNotFoundError:
            return SDKResult(
                success=False,
                output="",
                error="claude CLI를 찾을 수 없습니다"
            )
        except Exception as e:
            return SDKResult(
                success=False,
                output="",
                error=str(e)
            )

    def is_available(self) -> bool:
        """SDK 사용 가능 여부 확인"""
        try:
            result = subprocess.run(
                ['claude', '--version'],
                capture_output=True,
                text=True,
                timeout=5
            )
            return result.returncode == 0
        except Exception:
            return False
