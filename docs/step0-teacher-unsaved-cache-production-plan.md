# Step0 teacher unsaved cache 本番反映計画

## 文書の位置づけ

この文書は、Step0-A/B/Cで実装・検証した teacher unsaved cache 高速化を本番へ反映する際の手順、判定基準、確認項目、rollback方針を定める。

現時点では本番反映を実施しない。本番Apps Script、本番DB、本番トリガーには触れず、実施日時・担当者・承認者を決めたうえで別途この手順を実行する。

## 1. Step0の目的

教師画面の初期表示で行っていた過去未保存授業の都度集計を、事前構築したキャッシュの参照へ切り替える。これにより、通常の画面表示では重い旧APIを呼ばず、以下を実現する。

- 教師単位の未保存件数と詳細をFast APIから取得する。
- キャッシュが利用できない場合は0件と誤表示せず、状態を画面に表示する。
- 保存直後は画面上の件数をローカル減算し、次回buildで永続状態と再同期する。
- キャッシュを時間主導トリガーで定期更新できるようにする。
- 旧APIは比較検証とrollbackのため残す。Fast APIから旧APIへの自動フォールバックは行わない。

Step0ではUIの大改修、ダッシュボード化、一括保存、2コマ連続表示、詳細の本格的なページングは行わない。

## 2. Step0-A/B/Cの検証結果

すべて検証用Apps Script・検証用DBで確認済み。本番での確認は未実施である。

### Step0-A: キャッシュ構築とFast API

- `validateTeacherUnsavedCacheSheets()`：成功
- `debugBuildTeacherUnsavedCachePreview()`：成功
- `rebuildTeacherUnsavedSummaryCache()`：成功
- `getTeacherUnsavedSummaryFast()`：成功、実行時間は約2.7～4.3秒
- `getTeacherUnsavedDetailsFast()`：成功、実行時間は約2.1秒
- `debugCompareTeacherUnsavedCacheForCurrentUser()`：成功、旧APIを含むため約23.8秒
- T018の初回比較結果：旧サマリー78件、旧詳細78件、Fastサマリー78件、Fast詳細78件
- `summaryCount`、`detailCount`、`displayKeys`、`saveKeys`はすべて一致
- 差分配列はすべて空

### Step0-B: teacher画面のFast API切り替え

- `html/teacher.html`から`getTeacherUnsavedSummaryFast()`を呼び出せることを確認
- 詳細表示から`getTeacherUnsavedDetailsFast({ limit: 50, offset: 0 })`を呼び出せることを確認
- 自動初期表示・通常詳細表示から旧APIを呼ばない構成に変更済み
- `checkedAt`表示、非`ready`時の状態表示、50件を超える場合の控えめな注記を実装済み
- 保存後ローカル減算でT018が78件から77件になることを確認
- 再build後も旧処理・Fast処理が77件で一致し、`displayKey`、`saveKey`に差分なし

### Step0-C: 定期buildトリガー

- `installTeacherUnsavedCacheRebuildTriggers()`により検証環境で4件のCLOCKトリガー作成に成功
- 実行時刻は`Asia/Tokyo`の6時台、12時台、17時台、21時台
- `runTeacherUnsavedSummaryCacheRebuildTrigger()`の手動実行に成功
- install/removeは対象ハンドラー名が完全一致するトリガーだけを操作する
- トリガーの作成・削除は公開関数を明示実行した場合だけ行われる

## 3. 本番反映対象

### ソース

- `src/teacherUnsavedCacheService.js`
  - キャッシュスキーマ検証
  - キャッシュpreview/build
  - Fastサマリー・詳細API
  - 旧処理との比較用debug関数
  - rebuildトリガーの実行・状態確認・install・remove関数
- `html/teacher.html`
  - サマリーと詳細のFast API切り替え
  - 非`ready`状態と`checkedAt`の表示
  - 詳細先頭50件表示
  - 保存後ローカル減算のFastレスポンス対応

### 本番Operationスプレッドシート

- `teacherUnsavedSummaryCache`シート
- `teacherUnsavedDetailCache`シート

### 本番Apps Scriptの実行時設定

- `runTeacherUnsavedSummaryCacheRebuildTrigger`を呼ぶ時間主導トリガー4件
- トリガーは手動buildと本番画面確認が完了した後、指定した運用担当アカウントで作成する

## 4. 本番反映しない対象

- `src/config.js`の変更
- `appsscript.json`の変更
- Masterスプレッドシートのシート追加・スキーマ変更
- 既存Operationシートのスキーマ変更
- 既存の`getTeacherUnsavedSummary()`、`getTeacherUnsavedDetails()`の削除
- Fast APIから旧APIへの自動フォールバック
- teacher画面の大規模UI変更、ダッシュボード化、一括保存、2コマ連続表示
- 詳細の追加読み込みボタンや本格的なページング
- 検証環境のトリガーID、キャッシュデータ、実行履歴の本番へのコピー
- teacher unsaved cache以外の既存トリガーの変更・削除

debug関数はサービスファイルの一部として反映されるが、定期実行対象にはしない。承認された本番確認時だけ手動実行する。

## 5. 本番Operationに必要な追加シートとヘッダー

シート名は大文字・小文字を含めて完全一致させる。1行目のヘッダーは、列数、名前、順序が以下と完全一致する必要がある。余分な列を追加しない。作成時は2行目以降を空にし、数式や手入力データを置かない。

### teacherUnsavedSummaryCache

| 列 | ヘッダー |
|---:|---|
| A | `snapshotId` |
| B | `cacheDate` |
| C | `teacherId` |
| D | `teacherName` |
| E | `teacherEmail` |
| F | `startYmd` |
| G | `endYmd` |
| H | `unsavedCount` |
| I | `firstDate` |
| J | `lastDate` |
| K | `detailStartRow` |
| L | `detailCount` |
| M | `checkedAt` |
| N | `status` |
| O | `errorMessage` |

1行目に貼り付けるタブ区切り値：

```text
snapshotId	cacheDate	teacherId	teacherName	teacherEmail	startYmd	endYmd	unsavedCount	firstDate	lastDate	detailStartRow	detailCount	checkedAt	status	errorMessage
```

### teacherUnsavedDetailCache

| 列 | ヘッダー |
|---:|---|
| A | `snapshotId` |
| B | `cacheDate` |
| C | `teacherId` |
| D | `teacherName` |
| E | `teacherEmail` |
| F | `date` |
| G | `period` |
| H | `classId` |
| I | `sessionNumber` |
| J | `subjectName` |
| K | `targetLabel` |
| L | `isExperiment` |
| M | `displayKey` |
| N | `saveKey` |
| O | `sortKey` |
| P | `checkedAt` |

1行目に貼り付けるタブ区切り値：

```text
snapshotId	cacheDate	teacherId	teacherName	teacherEmail	date	period	classId	sessionNumber	subjectName	targetLabel	isExperiment	displayKey	saveKey	sortKey	checkedAt
```

キャッシュbuildが書き込むのはこの2シートの2行目以降である。既存の授業・出席・教員・時間割データは読み取り元であり、キャッシュbuildの書き込み先にはしない。

## 6. 本番反映手順

### 6.1 事前準備

1. 本番反映日時、実施者、確認者、rollback判断者を決める。
2. トリガー作成者となる運用担当アカウントを1つに決める。トリガーは作成者の権限で実行されるため、継続運用できるアカウントを使う。
3. 現在の本番デプロイバージョン、反映前コミット、既存トリガー一覧を記録する。
4. teacher画面で代表教員の旧API件数を記録する。最低限、未保存あり、未保存0件、実験・チーム担当を含む教員を選ぶ。
5. 本番OperationとMasterの対象シートが正常に参照できること、実施者が本番Operationの2つの新規シートを編集できることを確認する。

### 6.2 キャッシュシート準備

1. 本番Operationに`teacherUnsavedSummaryCache`を追加する。
2. 本番Operationに`teacherUnsavedDetailCache`を追加する。
3. それぞれの1行目へ「5. 本番Operationに必要な追加シートとヘッダー」の値を正確に設定する。
4. 2行目以降が空であること、余分な列がないことを確認する。

### 6.3 サーバーコード反映と手動build

1. 承認済みの通常リリース手段で`src/teacherUnsavedCacheService.js`を本番Apps ScriptのHEADへ反映する。この文書作成作業では`clasp push / pull`を実行しない。
2. 本番Webアプリの公開デプロイはまだ切り替えない。
3. `getTeacherUnsavedCacheSchema()`を実行し、想定シート名とヘッダーを確認する。
4. `validateTeacherUnsavedCacheSheets()`を実行し、両シートが`ok: true`であることを確認する。
5. `debugBuildTeacherUnsavedCachePreview()`を実行し、`wroteSheets: false`、件数、対象期間、warningを確認する。
6. previewに問題がなければ`rebuildTeacherUnsavedSummaryCache()`を1回だけ手動実行する。
7. build結果が`ok: true`、`wroteSheets: true`であり、`snapshotId`、`checkedAt`、`summaryRowCount`、`detailRowCount`が妥当であることを確認する。
8. `debugLogTeacherUnsavedSummaryFast()`と`debugLogTeacherUnsavedDetailsFast()`を実行し、`status: ready`、当日キャッシュ、同一`snapshotId`を確認する。
9. 代表教員で`debugLogCompareTeacherUnsavedCacheForCurrentUser()`を実行し、旧処理とFast処理の件数、`displayKey`、`saveKey`が一致することを確認する。

この段階で不一致、説明できないwarning、タイムアウト、権限エラーがある場合は公開デプロイへ進まず、トリガーも作成しない。

### 6.4 teacher画面反映と確認

1. `html/teacher.html`を含む本番Webアプリの新しいバージョンを作成し、承認済み手順で本番デプロイを更新する。
2. キャッシュbuild済みであることを再確認してからteacher画面を開く。
3. 「7. 本番確認手順」の画面確認を実施する。
4. 画面またはFast APIに問題があれば、トリガーを作成せずrollbackする。

### 6.5 トリガー作成（最後に実施）

> **必須順序:** トリガー作成は、手動build成功、Fast API比較成功、本番teacher画面確認成功のすべてを満たした後に行う。

1. トリガー作成者のアカウントで`debugLogTeacherUnsavedCacheRebuildTriggerStatus()`を実行し、既存対象トリガー数を記録する。
2. 同じアカウントで`installTeacherUnsavedCacheRebuildTriggers()`を明示実行する。初回はScriptAppとスプレッドシートへの権限承認が必要になる可能性がある。
3. 戻り値で`createdCount: 4`、`triggerCount: 4`、`configuredHours: [6, 12, 17, 21]`、`timeZone: Asia/Tokyo`を確認する。
4. `debugLogTeacherUnsavedCacheRebuildTriggerStatus()`とApps Scriptのトリガー管理画面で、対象ハンドラーのCLOCKトリガーが4件であることを確認する。
5. teacher unsaved cache以外の既存トリガーが変更されていないことを、事前記録と照合する。
6. 最初の定期実行後にApps Scriptの実行履歴とキャッシュの`checkedAt`を確認する。

時間主導トリガーは指定時刻ちょうどではなく、指定した時台の中で実行される。実装では`appsscript.json`のタイムゾーンに依存せず、`Asia/Tokyo`を明示している。

## 7. 本番確認手順

### 7.1 キャッシュとAPI

- [ ] `validateTeacherUnsavedCacheSheets()`が両シートとも`ok: true`
- [ ] 手動buildが`ok: true`かつ`wroteSheets: true`
- [ ] `cacheDate`が日本時間の当日
- [ ] `startYmd`と`endYmd`が想定対象期間
- [ ] サマリーと詳細の`snapshotId`が一致
- [ ] サマリー行の`status`が`ready`
- [ ] `checkedAt`が手動build時刻と整合
- [ ] `summaryRowCount`、`detailRowCount`、source row countsが妥当
- [ ] warningが0件、または全件について理由と影響を説明できる
- [ ] 代表教員で旧サマリー件数、旧詳細件数、Fastサマリー件数、Fast詳細件数が一致
- [ ] `displayKey`と`saveKey`の差分がない
- [ ] 未保存0件の教員でもFast結果が`ready`かつ0件

### 7.2 teacher画面

- [ ] 本日の授業一覧が従来どおり表示される
- [ ] 既存の7秒後＋800msの後読みタイミングで未保存サマリーが表示される
- [ ] Apps Script実行履歴で通常画面がFast APIを呼び、旧APIを呼んでいない
- [ ] 未保存件数、対象期間、`最終確認`が正しい
- [ ] 詳細を開くと先頭50件まで表示される
- [ ] 51件以上の場合に全件数と追加表示が必要な旨の注記が出る
- [ ] 「この授業を開く」から対象授業へ移動できる
- [ ] 過去の未保存授業を保存すると画面上の件数が1件減る
- [ ] 保存後の手動再buildで旧処理とFast処理が再び一致する
- [ ] `missing`、`stale`、`building`、`unavailable`、`error`が0件表示にならず、状態メッセージになる
- [ ] 出席入力、保存、過去授業検索など既存機能に回帰がない

### 7.3 トリガー

- [ ] 対象ハンドラーのトリガーが4件だけ存在する
- [ ] 4件とも`eventType: CLOCK`、`triggerSource: CLOCK`
- [ ] 設定時刻が日本時間の6時台、12時台、17時台、21時台
- [ ] 他ハンドラーのトリガーが事前記録どおり残っている
- [ ] 最初の定期実行が成功し、結果がログに出ている
- [ ] 少なくとも1日分の4回がタイムアウト・権限・クォータエラーなしで完了する
- [ ] 各実行後にキャッシュの`checkedAt`が更新され、Fast APIが`ready`を返す

## 8. rollback方針

### 基本原則

- キャッシュ2シートは派生データであり、既存の授業・出席データを正本として扱う。
- Fast APIは旧APIへ自動フォールバックしない。長時間Fast APIを利用できない場合は、teacher画面のデプロイを旧API版へ戻す。
- 緊急時も全トリガーを一括削除しない。対象ハンドラーだけを削除する。
- 調査用にログとキャッシュ内容を保持し、キャッシュシートを即時削除しない。

### 公開前に問題が見つかった場合

1. teacher画面の本番デプロイを更新しない。
2. `installTeacherUnsavedCacheRebuildTriggers()`を実行しない。
3. 手動buildで書かれた2つのキャッシュシートは調査用に保持する。
4. 原因を解消してpreview、手動build、比較を最初からやり直す。

### 公開後に問題が見つかった場合

1. トリガー作成済みなら、作成者のアカウントで`removeTeacherUnsavedCacheRebuildTriggers()`を実行する。
2. `deletedCount`と`remainingCount: 0`を確認し、他トリガーが残っていることを確認する。
3. 本番Webアプリのデプロイを、記録済みのStep0-B反映前バージョンへ戻す。
4. teacher画面が旧`getTeacherUnsavedSummary()`、`getTeacherUnsavedDetails()`で動作することを確認する。
5. `src/teacherUnsavedCacheService.js`のコード削除は急がない。トリガーと画面参照がなくなったことを確認後、別変更として判断する。
6. 2つのキャッシュシートは原則保持する。削除する場合は、トリガー停止、UI rollback、調査完了、承認の後にバックアップを取得して実施する。

コードrollbackを先に行いremove関数が利用できなくなった場合は、Apps Scriptのトリガー管理画面で`runTeacherUnsavedSummaryCacheRebuildTrigger`だけを手動削除する。

### 障害別の初動

| 状況 | 初動 |
|---|---|
| 手動build失敗・比較不一致 | 公開とトリガー作成を中止する |
| teacher画面のみ不具合 | 本番Webアプリを直前バージョンへ戻す |
| 定期トリガーのみ失敗 | 対象トリガーを停止し、原因調査中は必要に応じて手動buildする |
| キャッシュ破損・件数不一致 | 対象トリガー停止、画面rollback、ログと2シートを保全する |
| 権限・クォータエラー | 対象トリガー停止、作成者・権限・実行時間を確認する |

## 9. Step1 UI改善へ進む条件

次のすべてを満たした後にStep1へ進む。

- [ ] 本番手動build、Fast API、旧API比較が成功している
- [ ] 代表教員だけでなく、未保存0件、実験、チーム担当を含む確認が完了している
- [ ] teacher画面の初期表示、詳細表示、授業移動、保存後減算に問題がない
- [ ] 保存後の再buildで旧処理とFast処理が一致する
- [ ] 非`ready`状態が0件と誤表示されない
- [ ] 4つの定期トリガーが少なくとも1日分すべて成功している
- [ ] キャッシュの鮮度、実行時間、警告、Apps Scriptクォータに運用上の問題がない
- [ ] rollback手順、実施者、直前デプロイバージョンが記録されている
- [ ] Step0に関する未解決の重大・高優先度不具合がない

Step1では、初期表示の即時化、詳細ページング、表示情報の整理などを別変更として計画する。Step0の安定性確認と同時に大規模UI変更を混在させない。

## 10. 運用上の注意点

- **トリガー作成は必ず、手動buildと本番teacher画面確認の後に行う。**
- install/remove/statusの対象は、実行ユーザーが現在のプロジェクトに作成したトリガーである。作成者を統一し、別アカウントによる重複作成を避ける。
- `installTeacherUnsavedCacheRebuildTriggers()`は同名ハンドラーの既存トリガーを削除して4件を再作成する。他ハンドラーは削除しない。
- トリガー初回作成時はScriptApp権限、定期build実行には対象スプレッドシートへの権限が必要になる可能性がある。
- トリガーは日本時間の各時台で動作し、指定時刻ちょうどの実行は保証されない。
- Fast APIが`missing`、`stale`、`building`、`unavailable`、`error`の場合、teacher画面は旧APIへフォールバックしない。
- `checkedAt`と`cacheDate`を監視し、トリガーが存在するだけで正常と判断しない。
- キャッシュシートを手編集しない。修復は原則として原因を解消したうえで再buildする。
- 実行ログに教員情報や差分キーが含まれる可能性があるため、ログ共有範囲と保持期間に注意する。
