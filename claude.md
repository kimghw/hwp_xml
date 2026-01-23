# 주의: 사용자가 입력을 요청한 경우에만 수정하세요.
# 간략히 요점만 전체적인 맥락만 참고할 수 있도록 작성해 주세요.

## WSL에서 Windows Python 실행

```bash
cmd.exe /c "cd /d C:\hwp_xml\win32 && python insert_ctrl_id.py" 2>&1
```

- `cmd.exe /c` : Windows cmd 실행
- `cd /d C:\경로` : Windows 경로로 이동
- `2>&1` : stderr도 출력


## check_, test_ 접미사는 삭제해도 되며, 가능한 테스트 파일이나 체크 파일에 메인 로직을 넣지 말고 기존 로직을 쓰고, 기존 로직을 수정하고 핵라  