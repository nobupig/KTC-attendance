# Step0＋Step1-B 本番反映前 readiness checklist

## 1. 文書の目的

この文書は、Step0の教員未保存キャッシュFast化と、Step1-Bの科目担当画面授業カードUI改修を本番へ反映する前に、反映可否を判断するためのreadiness checklistである。

本番反映の具体的な操作手順書ではない。実施順序の原則、必須条件、中止条件、rollback準備が揃っているかを確認し、すべての必須項目を満たした場合だけ本番反映作業へ進む。

このStepでは本番DB、本番Apps Script、本番トリガーに触れない。`clasp push / pull`、Apps Script実行、DB操作、トリガー作成・削除も行わない。

詳細な本番反映手順は[Step0 teacher unsaved cache 本番反映計画](./step0-teacher-unsaved-cache-production-plan.md)、Step1-Bの変更・検証内容は[Step1-B 科目担当画面 授業カードUI検証・本番反映準備メモ](./step1-b-teacher-card-ui-verification.md)を参照する。

### 現時点のreadiness判断

**No-Go（確認継続）**

検証環境の6時台・12時台・17時台・21時台rebuild triggerと主要な画面検証は成功している。ただし、本番Apps Script ID、本番Operation ID、本番Master ID、production Operationのcache sheets準備、反映前の本番デプロイバージョンとrollback対象の記録、代表確認データと担当者の決定が未完了である。未チェックの必須項目が1件でもある間は、本番反映を開始しない。

## 2. 対象変更

### Step0：教員未保存キャッシュFast化

- `src/teacherUnsavedCacheService.js`のteacher unsaved cache service
- production Operationの`teacherUnsavedSummaryCache`
- production Operationの`teacherUnsavedDetailCache`
- `getTeacherUnsavedSummaryFast()`
- `getTeacherUnsavedDetailsFast({ limit, offset })`
- `rebuildTeacherUnsavedSummaryCache()`
- `runTeacherUnsavedSummaryCacheRebuildTrigger()`
- rebuild用時間主導トリガー
- `html/teacher.html`の未保存サマリー・詳細取得先のFast API切り替え

### Step1-B：科目担当画面の授業カードUI

- 授業カードの`class-card`、`class-card-top`、`period-badge`、`active`を含むDOM構造
- 保存済み、未保存、一部保存済み、確認中、確認失敗の状態pill
- 保存状態確認中・確認失敗時の`欠席者なしで保存`安全ガード
- 操作ボタン周辺の案内文、保存結果メッセージ、`title`
- 本日・検索日の担当授業件数概要
- 保存状態別chip
- 状態別カードclassと未保存カードの控えめな強調
- 768px以下の概要chip・カード・操作ボタン表示

Step1-Bでは保存API、Fast API、DB構造、保存データ仕様を変更していない。

## 3. 本番反映前の必須条件

この章の未チェック項目がすべて完了し、証跡を記録するまで本番反映は`No-Go`とする。

### 3.1 リポジトリと反映対象

- [ ] Step0、Step1-B、readiness documentを含む最新コミットがGitHubへpush済み
- [ ] ローカルHEADとGitHub上の対象コミットが一致している
- [ ] 作業ツリーがcleanで、未追跡・未コミット・未pushのファイルがない
- [ ] 本番反映対象のcommit hashを記録した
- [ ] Step2・Step3など対象外の変更が混在していない
- [ ] `git diff --check`が成功している

### 3.2 検証Apps Script・検証Operation

- [x] 検証Apps ScriptにStep0-A〜Cが反映されている
- [x] 検証Apps ScriptにStep1-B-0〜B-4が反映されている
- [x] 検証Operationに`teacherUnsavedSummaryCache`が存在する
- [x] 検証Operationに`teacherUnsavedDetailCache`が存在する
- [x] `validateTeacherUnsavedCacheSheets()`が成功している
- [x] read-only previewが成功している
- [x] 検証環境でmanual rebuildが成功している
- [x] `getTeacherUnsavedSummaryFast()`が`ready`を返す
- [x] `getTeacherUnsavedDetailsFast()`が正常に詳細を返す
- [x] 旧処理とFast処理の件数、`displayKey`、`saveKey`に差分がない
- [x] 保存後ローカル減算と再build後の再一致を確認している

### 3.3 rebuild triggerの検証状況

- [x] 6時台トリガーが成功している
- [x] 12時台トリガーが成功している
- [x] 17時台トリガーが成功している
- [x] 21時台トリガーが成功している
- [x] 4回すべてでタイムアウト、権限、クォータ、Spreadsheetアクセスエラーがない
- [x] 各実行後に`checkedAt`と`snapshotId`が更新され、Fast APIが`ready`を返す

検証環境のtrigger確認結果：

| 時間帯 | `ok` | `checkedAt` | `teacherCount` | `summaryRowCount` | `detailRowCount` | `warningCount` | `elapsedMs` |
|---|---:|---|---:|---:|---:|---:|---:|
| 6時台 | `true` | `2026-08-16 06:50:56` | 95 | 95 | 5052 | 0 | 57325 |
| 12時台 | `true` | `2026-08-16 12:14:29` | 95 | 95 | 5051 | 0 | 41342 |
| 17時台 | `true` | `2026-08-16 17:22:46` | 95 | 95 | 5051 | 0 | 93175 |
| 21時台 | `true` | `2026-08-16 21:43:45` | 95 | 95 | 5051 | 0 | 34081 |

12時台の`snapshotId`は`20260816_121429_367__44ee5861-0e86-4052-b1ba-2363397d99d3`である。

17時台の`elapsedMs`は93175msで、実行は成功し、`warningCount`は0だった。ただし、他の時間帯より長いため、本番反映時も`elapsedMs`を記録して観察する。21時台は34081msに戻っているため、現時点では継続悪化とは判断しない。

17時台build後の`getTeacherUnsavedSummaryFast()`確認結果：

| 項目 | 結果 |
|---|---|
| status | `ready` |
| `teacherId` | `T018` |
| `teacherName` | 安井宣仁 |
| `unsavedCount` | 76 |
| `detailCount` | 76 |
| `checkedAt` | `2026-08-16 17:22:46` |
| `cacheDate` | `2026-08-16` |
| `snapshotId` | `20260816_172246_354__6a534c32-232d-435f-ade5-5774e9b7574d` |

検証環境のStep0-C trigger安定確認は完了した。ただし、本番反映のGo条件はまだ満たしていない。次に確認するのは、本番側ID、production Operationのcache sheets準備、反映前の本番デプロイバージョンとrollback対象、代表確認データと担当者である。

### 3.4 科目担当画面

- [x] 本日の授業0件で従来どおり空表示になる
- [x] 本日の授業ありでカード一覧と概要が表示される
- [x] 過去授業`2026/06/24`を検索できる
- [x] 過去授業カードが2件表示される
- [x] `検索日の担当授業：2件`が表示される
- [x] 保存状態別chipが表示される
- [x] 未保存カードとactiveカードを識別できる
- [x] 過去未保存サマリーがFast APIで表示される
- [x] `欠席者を入力`から名簿を開ける
- [x] `欠席者なしで保存`が正常に動作する
- [x] 保存状態確認中はクイック保存だけがdisabledになる
- [x] 工学実験の集約カード、班選択、保存導線が維持されている
- [x] 360px相当の幅で概要chip、カード、操作ボタンが破綻しない

### 3.5 長時間実行への備え

- [x] Step1-B-4再実装後、本日授業・過去授業が正常表示されることを再確認した
- [x] 長時間実行を本番反映の中止条件として明文化した
- [x] 発生時に関数名、開始時刻、実行時間、終了状態、エラー内容を記録する方針がある
- [ ] 17時台・21時台を含む追加監視で長時間実行が再発していない
- [ ] 長時間実行が再発した場合の判断者とrollback担当者を決めた

## 4. 本番側で事前に準備するもの

この章は本番反映作業を開始する前に、値と担当者を別の承認済み作業記録へ記入する。本番IDやURLをこのリポジトリへ不用意に記録しない。

### 4.1 production Operationのキャッシュシート

- [ ] production Operationに`teacherUnsavedSummaryCache`を作成した
- [ ] production Operationに`teacherUnsavedDetailCache`を作成した
- [ ] 両シートの1行目が正確なヘッダーで、列数・名前・順序が一致する
- [ ] 両シートの2行目以降が空で、数式や手入力データがない

`teacherUnsavedSummaryCache`のヘッダー：

```text
snapshotId	cacheDate	teacherId	teacherName	teacherEmail	startYmd	endYmd	unsavedCount	firstDate	lastDate	detailStartRow	detailCount	checkedAt	status	errorMessage
```

`teacherUnsavedDetailCache`のヘッダー：

```text
snapshotId	cacheDate	teacherId	teacherName	teacherEmail	date	period	classId	sessionNumber	subjectName	targetLabel	isExperiment	displayKey	saveKey	sortKey	checkedAt
```

### 4.2 本番の接続先とrollback情報

- [ ] 本番Apps Script IDを確認し、検証Apps Script IDと異なることを照合した
- [ ] 本番Operation IDを確認し、検証Operation IDと異なることを照合した
- [ ] 本番Master IDを確認し、検証Master IDと異なることを照合した
- [ ] `.clasp.json`が本番反映対象のApps Scriptを向いていることを複数人で確認した
- [ ] `src/config.js`の本番Operation ID・本番Master IDを変更せず確認した
- [ ] 反映前の本番WebアプリURLを記録した
- [ ] 反映前の本番デプロイバージョンを記録した
- [ ] rollback対象のWebアプリバージョンとcommit hashを記録した
- [ ] 本番トリガーの既存一覧と作成者を記録した

### 4.3 代表確認データと担当者

- [ ] 本日授業ありの代表教員を決めた
- [ ] 本日授業なしの代表教員を決めた
- [ ] 過去授業検索の代表日・代表授業を決めた
- [ ] 未保存あり・未保存0件の代表教員を決めた
- [ ] 工学実験を含む代表教員・日付・授業を決めた
- [ ] 通常保存と`欠席者なしで保存`に使用する確認対象授業を決めた
- [ ] 反映実施者、画面確認者、Apps Script実行履歴確認者、rollback判断者を決めた

## 5. 本番反映順序の原則

この章は操作手順の詳細ではなく、依存関係を壊さないための順序原則である。各操作の詳細、権限、確認値はStep0本番反映計画に従う。

1. GitHub上の対象コミット、cleanな作業ツリー、本番ID、rollback対象を確定する。
2. Step0のサーバー側コードを先に本番Apps Scriptへ反映する。
3. production Operationへ2つのキャッシュシートと正確なヘッダーを準備する。
4. 本番で`validateTeacherUnsavedCacheSheets()`を実行する。
5. 本番でread-only previewを実行し、書き込みなしで対象件数とwarningを確認する。
6. preview成功後にmanual rebuildを1回実行する。
7. Fastサマリー、Fast詳細、代表教員の旧処理比較を確認する。
8. Step0を利用する科目担当画面を確認する。
9. Step0が安定してからStep1-BのUIを反映する。
10. Step1-B反映直後の画面・保存・実行履歴を確認する。
11. 手動build、Fast API、画面確認のすべてが成功した後、最後にrebuild triggerを作成する。
12. Step2の一括保存とStep3の2コマ連続表示は、この反映へ混在させない。

**trigger作成をmanual rebuildや画面確認より先に行わない。** UIに問題がある状態、Fast APIが非`ready`の状態、接続先が不明な状態ではtriggerを作成しない。

## 6. 本番反映直後の確認

### 6.1 Apps Script・Fast API

- [ ] `getTeacherUnsavedSummaryFast()`が`ok: true`かつ`status: ready`
- [ ] `getTeacherUnsavedSummaryFast()`の`cacheDate`と`checkedAt`が妥当
- [ ] `getTeacherUnsavedDetailsFast()`が同じsnapshotの詳細を返す
- [ ] `getTeacherUnsavedDetailsFast({ limit: 50, offset: 0 })`の件数と`hasMore`が妥当
- [ ] `getTodayClassesForCurrentUser()`が正常終了し、実行時間に異常がない
- [ ] `getClassesForCurrentUserByDate()`が正常終了し、実行時間に異常がない
- [ ] `getSaveStatusForTeacherSessions()`が正常終了し、対象カードの状態が更新される
- [ ] Apps Script実行履歴にSpreadsheetアクセス、権限、タイムアウト、クォータエラーがない

### 6.2 科目担当画面

- [ ] 本日の授業ありで、担当授業件数とカード件数が一致する
- [ ] 本日の授業なしで、概要を出さず空表示になる
- [ ] 保存状態が`確認中`から保存済みまたは未保存へ更新される
- [ ] 過去授業検索が成功し、検索日の担当授業件数とカード件数が一致する
- [ ] 未保存サマリー、詳細、最終確認時刻が正しい
- [ ] `欠席者を入力`から名簿を開ける
- [ ] 通常の出席内容を保存できる
- [ ] 未保存授業を`欠席者なしで保存`できる
- [ ] 保存状態確認中・確認失敗時の安全ガードが働く
- [ ] 工学実験が集約カード1件・班選択ボタン1件として表示される
- [ ] 工学実験で班選択後に名簿を開き、保存できる
- [ ] 360px相当の幅で概要chip、カード上段、操作ボタン、右側入力欄が破綻しない

### 6.3 反映直後の記録

- [ ] 本番デプロイバージョンとcommit hashを記録した
- [ ] manual rebuildの`snapshotId`、件数、`warningCount`、`elapsedMs`を記録した
- [ ] 主要5関数の実行時間を記録した
- [ ] 代表教員・日付・授業ごとの画面確認結果を記録した
- [ ] trigger作成前のGo／No-Go判断を記録した

## 7. 中止条件

次のいずれかが発生した場合は`No-Go`とし、その段階以降の反映、画面操作、trigger作成を中止する。原因と影響範囲を確認できるまで再開しない。

- [ ] Spreadsheetアクセスエラーが発生した
- [ ] `getTodayClassesForCurrentUser()`が通常より著しく長時間実行、タイムアウト、または未完了になった
- [ ] `getClassesForCurrentUserByDate()`が通常より著しく長時間実行、タイムアウト、または未完了になった
- [ ] `getTeacherUnsavedSummaryFast()`が`missing`、`stale`、`unavailable`、`error`を返した
- [ ] キャッシュシートが存在しない、またはヘッダーの列数・名前・順序が一致しない
- [ ] previewまたはmanual rebuildの`warningCount > 0`
- [ ] manual rebuildが`ok: true`にならない、または件数を説明できない
- [ ] 保存状態が`確認中`から更新されない
- [ ] 通常保存または`欠席者なしで保存`に失敗した
- [ ] 工学実験の集約カード、班選択、名簿、保存のいずれかが壊れた
- [ ] 本番と検証のApps Script、Operation、MasterのIDを取り違えた疑いがある
- [ ] `.clasp.json`の向き先が不明、未確認、または本番対象と一致しない
- [ ] 本番反映対象外のファイルやStep2・Step3の変更が混在している
- [ ] rollback対象バージョンが不明、またはrollback担当者へ連絡できない

このチェックリストでは、中止条件の項目は通常は未チェックのままとする。いずれかが該当した場合にチェックし、反映を止めた証跡として使用する。

## 8. rollback方針

### Step1-B

- `html/teacher.html`のStep1-B差分をrevertし、本番Webアプリを記録済みの直前バージョンへ戻す。
- DB、キャッシュシート、トリガー、保存データは変更しない。

### Step0

- Fast APIまたはteacher画面連携に問題がある場合は、Fast API利用前の`teacher.html`を含む本番Webアプリバージョンへ戻す選択肢を取る。
- キャッシュシートは派生データと調査証跡として保持し、緊急時に削除しない。
- rebuild triggerを作成済みで停止が必要な場合は、`removeTeacherUnsavedCacheRebuildTriggers()`を権限のある作成者が後から明示実行する。
- remove関数を利用できない場合は、Apps Script管理画面で`runTeacherUnsavedSummaryCacheRebuildTrigger`のトリガーだけを削除する。他のトリガーは削除しない。
- 既存の授業・出席・保存データは正本であり、DB保存データのrollbackは不要である。

### rollback後の確認

- [ ] 本日の授業一覧を取得できる
- [ ] 過去授業を検索できる
- [ ] `欠席者を入力`から名簿を開ける
- [ ] 通常保存が成功する
- [ ] `欠席者なしで保存`が成功する
- [ ] 工学実験の班選択と保存が成功する
- [ ] Apps Script実行履歴に新しいエラーがない
- [ ] rollback後のWebアプリバージョンと確認結果を記録した

## 9. Step2・Step3へ進む条件

次のすべてを満たすまで、Step2の一括保存API/UIとStep3の2コマ連続授業表示グルーピングへ進まない。

- [ ] Step0とStep1-Bの本番反映が完了している
- [ ] 本番反映直後の必須確認がすべて成功している
- [ ] 6時台、12時台、17時台、21時台の本番rebuild triggerが安定して成功している
- [ ] 各trigger後にFast APIが`ready`を返し、キャッシュの鮮度に問題がない
- [ ] 実利用で授業取得APIの長時間実行が再発していない
- [ ] 実利用で保存状態表示、通常保存、クイック保存に問題がない
- [ ] 工学実験の実利用導線に問題がない
- [ ] デスクトップとスマートフォンでStep1-BのUIに問題がない
- [ ] rollback対象バージョン、担当者、実施結果が記録されている
- [ ] Step0・Step1-Bに未解決の重大または高優先度不具合がない

条件を満たした後も、Step2とStep3はそれぞれ独立した設計・実装・検証・本番判断として進める。
