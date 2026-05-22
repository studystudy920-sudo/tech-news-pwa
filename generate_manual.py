from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()

# デフォルトフォント設定（日本語対応）
style = doc.styles['Normal']
style.font.name = 'Yu Gothic'
style.font.size = Pt(11)

# タイトル
title = doc.add_heading('クラウド版 Claude Code 利用マニュアル', level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph(
    'ウェブから起動した remote execution environment（クラウド版 Claude Code）でできることをまとめたマニュアルです。'
)

# 1. 環境の特徴
doc.add_heading('1. 環境の特徴', level=1)
features = [
    '隔離された一時コンテナで動作する。リポジトリは起動時にクローンされ、セッション終了後に破棄される。',
    '残したい変更は必ず commit & push する必要がある。',
    '作業ディレクトリの例: /home/user/tech-news-pwa（git リポジトリ）。',
    '開発ブランチ例: claude/relaxed-archimedes-Sf7EX',
    'ネットワーク到達範囲はユーザーが選択したネットワークポリシーに従う。',
]
for f in features:
    doc.add_paragraph(f, style='List Bullet')

# 2. できること
doc.add_heading('2. できること', level=1)

doc.add_heading('2.1 コード開発', level=2)
for item in [
    'ファイルの読み取り・編集・新規作成（Read / Edit / Write ツール）',
    'Bash コマンドの実行（テスト、リント、ビルド等）',
    'コードベース全体の探索（Agent の Explore など）',
    '複数ステップにわたる実装タスク（Agent の general-purpose）',
    '実装計画の作成（Agent の Plan）',
]:
    doc.add_paragraph(item, style='List Bullet')

doc.add_heading('2.2 GitHub 連携（MCP 経由）', level=2)
for item in [
    'PR の作成・更新・マージ',
    'Issue の参照・コメント・作成',
    'ブランチ、コミット、タグの操作',
    'PR レビューへの対応（コメント返信、コードの修正）',
    'CI イベントの購読（subscribe_pr_activity）による自動修正ループ',
    '対象リポジトリは studystudy920-sudo/tech-news-pwa に限定',
    'gh CLI は使用不可。すべて mcp__github__* ツールを使う。',
]:
    doc.add_paragraph(item, style='List Bullet')

doc.add_heading('2.3 外部リソース', level=2)
for item in [
    'WebFetch / WebSearch によるウェブ情報取得',
    'Notion 連携（ページ作成、検索、コメント、データベース操作 等）',
    'Google Drive 連携（ファイル作成、検索、メタデータ取得、コピー 等）',
    '一部 MCP ツールは認証が必要。',
]:
    doc.add_paragraph(item, style='List Bullet')

doc.add_heading('2.4 利用可能なスキル（/コマンド）', level=2)
skills = [
    ('/verify', '変更を実際に動かして検証する'),
    ('/code-review', '差分に対するコードレビューを実施する'),
    ('/security-review', 'セキュリティレビューを実施する'),
    ('/run', 'プロジェクトのアプリを起動して動作確認する'),
    ('/init', 'CLAUDE.md を生成する'),
    ('/review', 'PR をレビューする'),
    ('/loop', '定期実行（例: 5分ごと）'),
    ('/session-start-hook', 'Web セッション用の SessionStart hook を設定'),
    ('/update-config', 'settings.json による harness 設定変更'),
    ('/keybindings-help', 'キーバインドのカスタマイズ'),
    ('/fewer-permission-prompts', '権限プロンプトを減らす許可設定'),
    ('/claude-api', 'Claude API / Anthropic SDK 関連の開発支援'),
]
table = doc.add_table(rows=1, cols=2)
table.style = 'Light Grid Accent 1'
hdr = table.rows[0].cells
hdr[0].text = 'コマンド'
hdr[1].text = '内容'
for cmd, desc in skills:
    row = table.add_row().cells
    row[0].text = cmd
    row[1].text = desc

# 3. Git 運用ルール
doc.add_heading('3. Git 運用ルール', level=1)
for item in [
    '指定された開発ブランチで作業すること（例: claude/relaxed-archimedes-Sf7EX）。',
    'push は git push -u origin <branch-name> を使用。',
    'ネットワークエラー時は最大4回、指数バックオフ（2s, 4s, 8s, 16s）でリトライ。',
    'PR の作成はユーザーから明示的に依頼があった場合のみ行う。',
    'fetch / pull は対象ブランチを明示する。',
    'hooks のスキップ（--no-verify 等）はユーザーの明示的指示がない限り行わない。',
]:
    doc.add_paragraph(item, style='List Bullet')

# 4. PR アクティビティ監視
doc.add_heading('4. PR アクティビティ監視', level=1)
doc.add_paragraph(
    'subscribe_pr_activity で PR の CI / コメント / レビューを購読すると、'
    '<github-webhook-activity> としてイベントが届く。'
)
for item in [
    'CI 失敗の自動修正（ブランチへの push を含む）',
    'レビューコメントへの応答（曖昧な場合は AskUserQuestion で確認）',
    '不要な購読は unsubscribe_pr_activity で解除',
    'ポーリングや sleep ループは使わない（イベント駆動）',
]:
    doc.add_paragraph(item, style='List Bullet')

# 5. 注意事項
doc.add_heading('5. 注意事項・制限', level=1)
for item in [
    'コンテナはエフェメラルなため、コミットしない変更は失われる。',
    '対象 GitHub リポジトリ以外への操作は拒否される。',
    '機密情報（.env、credentials など）の commit は避ける。',
    '破壊的な git 操作（force push、reset --hard 等）はユーザーの明示的指示が必要。',
    '外部公開ツール（pastebin、図解ツール 等）に機密情報を送らない。',
]:
    doc.add_paragraph(item, style='List Bullet')

# 6. 参考リンク
doc.add_heading('6. 参考リンク', level=1)
doc.add_paragraph('公式ドキュメント: https://code.claude.com/docs/en/claude-code-on-the-web')

# 保存
output_path = '/home/user/tech-news-pwa/ClaudeCode_Cloud_Manual.docx'
doc.save(output_path)
print(f'Saved: {output_path}')
