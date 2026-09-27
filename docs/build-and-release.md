# ビルドとリリースの手引き

## 今回の到達点

リリース対象は `0.4.0`。この手引きのスクリプトは検証と梱包までを行い、公開操作は別に実施する。
リリース候補の成果物を作ったことと、公開済みであることは区別する。
利用者向けの変更点は[更新履歴](../CHANGELOG.md)、移行手順は[READMEの0.4.0への更新](../README.md#040への更新)にまとめる。

この文書は、ビルドまわりを後から復習するための読み順も兼ねる。
最初は次の関係だけ押さえればよい。

```text
利用アプリのcsproj（PackageReference + DocumentAsData項目）
    → NuGetがprops/targetsを自動import
    → 通常ビルドでGenerateDocumentAsDataタスクを実行
    → CodeGenerationがCore経由でDOCXを読み、.g.csを生成
    → 利用アプリが生成ソースをコンパイル
    → 実行時はCoreを通じてDOCXを読み書き
```

## どのファイルを読むか

| ファイル・要素 | 役割と、必要になった理由 |
| --- | --- |
| `global.json` | 開発・CIで使用する.NET SDKを揃える |
| `src/Directory.Build.props` | 4パッケージのバージョン・作者・MIT・README・リポジトリ情報を揃える。テストプロジェクトには適用しない |
| 各 `src/*/*.csproj` | Core、生成器、ビルドタスクそれぞれの依存関係と配布内容を定める |
| `buildTransitive/*.props` | 評価の早い段階で生成ファイルの既定項目への取り込みを防ぎ、ビルドアクション候補を登録する |
| `buildTransitive/*.targets` | 生成ソースのCompile登録、元DOCXとの紐付け、通常ビルド時の生成、Cleanを定める |
| `GenerateDocumentAsData.cs` | 文書と辞書を読み、診断と生成を呼び出す。生成器そのものとは責務を分ける |
| `scripts/Prepare-Release.ps1` | Releaseビルド・テスト・書式検査・梱包を順に実行する。公開はしない |
| `.github/workflows/verify.yml` | 上記スクリプトをWindowsのGitHub Actionsで実行し、結果をartifactとして保存する。公開用権限やトークンは使わない |

`props/targets` の「評価」と、タスクの「実行」は別である。
IDEがソースの一覧を知るだけのときにもCompile項目は必要だが、そのたびにDOCXを開いて生成する必要はない。
そのため `DesignTimeBuild=true` では既存の生成ソースだけを参照し、通常ビルドで変更を反映する。

## パッケージはどう分かれるか

* `Marimo.DocumentAsData.Core`：DOCXを読み書きする本体。
* `Marimo.DocumentAsData.CodeGeneration`：Coreに依存する生成器。
* `Marimo.DocumentAsData.Build`：Coreに依存し、ビルドへ組み込む設定とタスク、生成用DLLを配布する。
* `Marimo.DocumentAsData`：通常の利用者向けの統合パッケージ。Buildと、その依存先のCoreを導入する。

Coreだけでも通常のライブラリとして使える。0.4.0では生成器をBuildのtools内へ同梱し、アプリの実行時依存には渡さない。
生成APIをアプリから直接呼ぶ場合は、CodeGenerationパッケージを明示的に参照する。
統合パッケージ自身にはDLLやビルド設定を重複して梱包せず、Buildの`buildTransitive`設定を利用する。

Buildパッケージでは、タスク本体と実行に必要なDLLを `tools/net10.0/` へ梱包する。
利用アプリにMSBuildのタスク型まで公開する必要はないので、タスクDLLを `lib/` に入れない。
`IncludeBuildOutput=false` はこのための設定であり、「DLLを一切配布しない」という意味ではない。

`lib/net10.0/_._` は空ファイルである。
参照DLLを追加せずに.NET 10の依存関係と対象フレームワークを対応付け、NuGetのNU5128を解消する。
警告を抑制する設定ではない。[NuGet公式説明](https://learn.microsoft.com/en-us/nuget/reference/errors-and-warnings/nu5128)

## 初版のビルド連携をTDDで確認したこと

テスト本体は `tests/DocumentAsData.CodeGeneration.Tests/MSBuild連携タスクのテスト.cs`、
一時プロジェクト作成は同プロジェクトの `テスト補助/MSBuild連携テストプロジェクト.cs` にある。
テストはパッケージを実際に作り、ローカルフィードから別プロジェクトへ復元する。
リポジトリ内のDLLを直接参照するテストだけでは、同梱漏れや自動importの問題は確認できないためである。

| 確認対象 | 観測した失敗または結果 | 対応 |
| --- | --- | --- |
| タスクDLLがアプリ参照へ混入しない | コンパイラ参照にBuild.dllが含まれていた | `IncludeBuildOutput=false` |
| 3パッケージの公開情報 | 未設定の名前と既定バージョン1.0.0だった | 共通配布情報と役割別PackageId |
| 警告をエラーにした梱包 | BuildパッケージがNU5128で失敗 | 空のフレームワーク識別ファイル |
| Visual Studio版のMSBuild | 確認環境では既存実装のまま成功 | 本体は変更せず、テストホストの切り替えを用意 |
| Clean | 文書隣の生成ソースが残った。FileWritesへの登録だけでも残った | 現在の対象DOCXに対応する生成ソースだけをDeleteタスクへ渡す |
| パッケージからの実行 | 生成型でRead→Replace→SaveAs→再読込が成功 | 本体は変更せず、実行までの回帰テストを追加 |

Visual Studio版MSBuildとパッケージ実行の確認など、最初から通ったものをRedだったとは扱わない。
本体に新しい振る舞いを追加した箇所は、失敗を確認してから必要な実装を入れている。
初版の梱包整備では、コレクションや文書読み書きの公開APIは変更していない。0.4.0のAPI追加・変更はREADMEの更新案内を参照する。

標準Cleanは主に `bin`／`obj` の出力を片付けるため、文書隣の生成ファイルには別の指定が必要だった。
現在の対象から外された文書の古い生成ソースは、このCleanでも削除しない。既定Compileからの除外だけを行う。
生成物を複数プロジェクトで共有せず、手書きコードは別のpartialクラスへ置く。

## 開発環境と検証範囲

* .NET SDKは `global.json` の `10.0.302`（同一feature bandのパッチ更新を許可）。
* テスト・検証スクリプトはPowerShell 7を使用する。ライブラリ利用者にPowerShellは不要。
* Windowsでの全テストと、実際のパッケージ復元・ビルド・実行を検証する。
* Visual Studio版MSBuildは `18.8.2`（RuntimeType=Full）、dotnet版は `18.6.11`（Core）でパッケージ利用を確認した。
* Visual Studioの画面操作、古いMSBuild、他OS、Blazor WASM上での実行そのものは、この検証の保証範囲ではない。
* WordやLibreOfficeのインストールは自動テストに不要。

SpreadsheetAsDataのtargetsでは `UsingTask Runtime="NET"` も使用されているが、
このリポジトリの確認環境では追加しなくてもテストが成功したため、今回は本体へ移植していない。
他環境で問題が確認された場合に再現テストを起点に対応する。
`Runtime="NET"` 自体の適用条件は [MSBuild公式資料](https://learn.microsoft.com/en-us/visualstudio/msbuild/usingtask-element-msbuild?view=visualstudio) を参照する。

## Release候補を作成する

リポジトリのルートで次を実行する。

```powershell
pwsh -NoProfile -File ./scripts/Prepare-Release.ps1
```

初回のrestoreではNuGetから依存パッケージを取得する。パッケージ利用テストは外部フィードではなく、
その復元済みキャッシュとテスト内のローカルフィードを利用する。
そのテスト用プロジェクトでのみ `NuGetAudit=false` として、オフラインの復元検証と脆弱性監査を分けている。
通常のrestoreの監査は無効化していない。

必要なら、使用するVisual Studioの `MSBuild.exe` の実際のパスを指定する。

```powershell
pwsh -NoProfile -File ./scripts/Prepare-Release.ps1 -VisualStudioMSBuild '<MSBuild.exeの実際のパス>'
```

スクリプトは先にdotnet版で全テストを実行し、その後、指定したホストでパッケージ関連テストを再実行する。
常に新しい `artifacts/release/<実行ID>/` を作るため、以前の結果を削除したり混在させたりしない。

* `packages/`：4つの `.nupkg`
* `tests/`：Release全テストのTRX
* `visual-studio-tests/`：指定した場合の追加テスト結果
* `SHA256SUMS.txt`：配布物のSHA-256
* `VERIFICATION.md`：検証対象コミット、作業ツリー変更の有無、検証結果

途中で失敗した場合はそこで停止し、成功を示すVERIFICATION.mdは作らない。
成功したテストだけで、未実行環境も対応済みとは判断しない。
GitHub Actions上での実行結果はpush後に別途確認する。ローカル実行だけではクラウド側の成功は未確認である。

## 公開前にローカルパッケージを使う

利用側の `NuGet.Config` に、上で作成した `packages` ディレクトリをパッケージソースとして追加する。
既存の復元元を消す必要はない。NuGet.orgには外部依存の復元時に接続する。

```xml
<configuration>
  <packageSources>
    <add key="DocumentAsDataLocal" value="＜生成されたpackagesディレクトリの絶対パス＞" />
  </packageSources>
</configuration>
```

そのうえで、読み書きだけならCore、自動生成も使うならBuildをVersion="0.4.0"で参照する。
初めて試す場合は[READMEのコード生成入門](../README.md#はじめて使うコード生成でdocxを読み書きする)を参照する。
設定の詳細は[機能リファレンスのMSBuild連携](reference.md#msbuildで自動生成する)を参照する。
同じバージョンを再作成した際は、古いパッケージキャッシュを再利用していないことに注意する。
自動テストは復元先もテストごとに分離している。

## 公開直前のチェックリスト

1. PackageIdの所有権・利用可否、公開先、MITの著作権表示、最終バージョンを確認する。公開済みとは仮定しない。
2. バージョンを変える場合は `src/Directory.Build.props` と対応する配布検査の期待値・READMEを揃える。
3. `CHANGELOG.md` に前版からの追加機能・動作変更・修正を記録する。READMEの更新案内、利用例、対応範囲も実装と照合し、破壊的な変更や既知の制約を記載してコミットする。公開前の版は公開準備中と明記し、公開確認後に実際の公開日へ更新する。
4. クリーンな作業ツリーからPrepare-Releaseを再実行し、TRXとVERIFICATION.mdを確認する。
5. 4パッケージ、依存関係、README、MIT、SHA-256を確認する。未実行環境の保証を加えない。
6. GitHub Actions上の検証結果も確認する。公開権限・認証は別途用意する。
7. 承認したコミット・バージョン・成果物だけでタグ、GitHub Release、NuGet公開を行う。GitHub Releaseの説明にも該当バージョンの更新履歴を掲載する。

この準備スクリプトにはpush、タグ、Release作成、NuGet公開のコマンドを含めていない。
公開時は依存順（Core → CodeGeneration → Build → 統合パッケージ）と、NuGet側で各バージョンが利用可能になったことを確認する。
既に公開した同一バージョンは再公開せず、追加したパッケージだけを公開する。
公開後は新しい利用者プロジェクトから公開フィードだけで復元して、最小例を実行する。

## 初版の公開前整備で変更したファイル

既存の文書読み書きAPIの実装や固定DOCXは、この整備では変更していない。

* `src/Directory.Build.props`
* `src/DocumentAsData/DocumentAsData.csproj`
* `src/DocumentAsData.CodeGeneration/DocumentAsData.CodeGeneration.csproj`
* `src/DocumentAsData.Build/DocumentAsData.Build.csproj`
* `src/DocumentAsData.Build/PackagePlaceholder/_._`
* `src/DocumentAsData.Build/buildTransitive/Marimo.DocumentAsData.Build.targets`
* `tests/DocumentAsData.CodeGeneration.Tests/MSBuild連携タスクのテスト.cs`
* `tests/DocumentAsData.CodeGeneration.Tests/テスト補助/MSBuild連携テストプロジェクト.cs`
* `scripts/Prepare-Release.ps1`
* `.github/workflows/verify.yml`
* `README.md`
* `docs/design-v0.1.md`
* `docs/build-and-release.md`
