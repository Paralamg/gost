# Разработка

Проект использует [uv](https://docs.astral.sh/uv/).

```bash
git clone git@github.com:Paralamg/gost.git
cd gost
uv sync
uv run python examples/example.py
uv run pytest
```

## Тесты конвертера формул

Тесты конвертера LaTeX → OMML сверяются с эталоном Microsoft — `MML2OMML.XSL` из поставки
Office. Сам файл проприетарный, в репозиторий не входит и библиотекой не используется;
сверка нужна только при разработке и пропускается, если Office не установлен. Путь
задаётся переменной окружения `GOST_MML2OMML_XSL`.

## Публикация на PyPI

Сборка выполняется бэкендом `uv_build`, публикация — командой `uv publish`.

### 1. Подготовка

Перед выпуском поднимите версию в `pyproject.toml` (поле `version`) — PyPI не разрешает
повторно загрузить уже опубликованную версию, даже после удаления файла.

Получите API-токен: [pypi.org/manage/account/token](https://pypi.org/manage/account/token/)
(и отдельный — на [test.pypi.org](https://test.pypi.org/manage/account/token/) для тестовой площадки).
Токен начинается с `pypi-`. Передавайте его через переменную окружения, а не аргументом командной
строки — так он не попадёт в историю оболочки:

```powershell
# PowerShell
$env:UV_PUBLISH_TOKEN = "pypi-AgEIcHl..."
```

```bash
# bash
export UV_PUBLISH_TOKEN="pypi-AgEIcHl..."
```

### 2. Сборка

```bash
uv build
```

Артефакты появятся в `dist/`: `gost_docx-<версия>-py3-none-any.whl` и `gost_docx-<версия>.tar.gz`.
Если в `dist/` остались файлы прошлого выпуска, очистите каталог — `uv publish` загружает всё,
что там лежит:

```powershell
Remove-Item -Recurse -Force dist   # PowerShell
```

```bash
rm -rf dist                        # bash
```

### 3. Проверка на TestPyPI (рекомендуется)

```bash
uv publish --publish-url https://test.pypi.org/legacy/
```

Установка из TestPyPI — зависимости при этом берутся с основного PyPI:

```bash
pip install --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple/ gost-docx
```

### 4. Публикация

```bash
uv publish
```

Проверка установки из PyPI в чистом окружении:

```bash
uv run --with gost-docx --no-project -- python -c "from gost import WordBuilder; print(WordBuilder())"
```

### 5. Тег релиза

```bash
git tag v0.1.0
git push origin v0.1.0
```
