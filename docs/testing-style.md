# テストコードの変数方針

この文書は、DocumentAsData のテストでローカル変数と一時変数を扱う方針を定義する。

## var の使用方針

ローカル変数は原則として `var` を使用する。

```csharp
var rows = tested.ToArray();
var cell = sheet.Cells["A1"];
```

次の場合は具体型を明記してよい。

* フィールド宣言
* メソッドやプロパティの公開シグネチャ
* ラムダ式の代入先として型が必要な場合
* 具体型そのものをコンパイル時に確認するテスト
* `var` では意図が著しく分かりにくくなる場合

例外検証のラムダも原則として `var` で受ける。

```csharp
var action = () => tested.DoSomething();
```

デリゲート型そのものを示す必要がある場合や、戻り値を持つ処理、非同期処理を明示的に区別したい場合は、そのテストで必要な型を選ぶ。

共有するテスト対象のフィールドは、従来どおり具体型を明記する。

```csharp
readonly Workbook book;
readonly Table tested;
```

## 一時変数を作る基準

一時変数は、単に長い式へ英語名を付けるためには作らない。

次のいずれかに該当する場合に作る。

* 同じ値を複数回使用する
* 複数のAssertで同じ対象を検証する
* 列挙や評価を一度だけに固定する
* 例外アサーションを一度だけ実行して複数プロパティを検証する
* 式の一部が独立した概念を表す
* デバッグ時に途中結果を観察する価値が高い

```csharp
var rows = tested.ToArray();

rows
    .Select(row => row.Id)
    .Should().Equal(1, 2, 3);

rows
    .Select(row => row.Name)
    .Should().Equal("a", "b", "c");
```

この例では、同じ列挙結果を複数回検証し、列挙を一度だけに固定するため、一時変数に意味がある。

例外について複数の情報を確認する場合は、例外オブジェクトではなく例外アサーションを変数へ受ける。
`Which` は後続の `Should()` へつなぐために使い、変数へ受けない。

```csharp
var thrown = action
    .Should().Throw<TableMappingException>();

thrown.Which.TableName.Should().Be("Table1");
thrown.Which.ColumnName.Should().Be("string");
thrown.Which.PropertyName.Should().Be(nameof(TestRow.Value));
```

## FluentAssertions と null 検証

nullable な値を検証した後に同じ値を続けて使う場合は、まず `Should().NotBeNull()` で仕様として非 null を確認する。
FluentAssertions の `NotBeNull()` は nullable 解析に対応しているため、確認後の同じ変数は非 null として扱える。

```csharp
var property = type.GetProperty("MainRange");

property.Should().NotBeNull();
property.PropertyType.Should().Be(typeof(CellRange));
```

`NotBeNull()` で確認した値を使うためだけに、`!`、`?? throw`、`Which` を追加しない。

`Which` は、例外検証や型検証などで、FluentAssertions のチェーンとしてさらに検証を続ける場合に使う。nullable 解析を外すための一時変数化には使わない。

```csharp
action
    .Should().Throw<TableMappingException>()
    .Which.TableName
    .Should().Be("Table1");
```

## Fluentな検証句の改行

`Should()` と最終アサーションは、原則として同じ行に置く。
FluentAssertions は英文に近い検証句として読むため、`Should()` だけを単独行にして述語を分断しない。
`Which`、検証対象の短いプロパティ、`ToString()` なども、一続きの句として読める場合は無理に縦へ分割しない。

```csharp
actual.Should().Be(expected);

book.ReadTable<TestRow>("Table1")
    .Select(it => it.Value)
    .Should().Equal(1, 2, 3);

action
    .Should().Throw<TableMappingException>()
    .Which.TableName.Should().Be("Table1");
```

検証対象を作るチェーンは、声に出して読んだときの区切りや、意味の切れ目で改行してよい。
ただし、`Should()` の前で改行すること自体をルールにはしない。
検証対象の構築が長い場合など、読みやすくなるときの選択肢として扱う。
期待値が長い場合は、`Should()` とアサーション名ではなく、引数側を改行する。

```csharp
actual.Should().Equal(
    1,
    2,
    3);
```

## 一時変数を作らない基準

一度しか使用せず、右辺の式がそのまま意味を表している場合は、直接Assertへつなげる。

```csharp
book.ReadTable<IntegerOnlyRow>("Table1")
    .Select(row => row.IntegerValue)
    .Should().Equal(1, 2, 3);
```

次のような、一度しか使わない中間変数は原則として作らない。

```csharp
var actual = tested.Value;

actual.Should().Be(expected);
```

次の形を優先する。

```csharp
tested.Value.Should().Be(expected);
```

同様に、一度しか使わない `result` も原則として不要とする。

```csharp
book.ReadTable<TestRow>("Table1")
    .Should().BeEmpty();
```

特に、変数名が次のような一般名でしかなく、右辺以上の情報を加えていない場合はインライン化する。

```text
actual
result
value
data
item
object
```

ただし、その変数を複数回使う場合や、評価回数を固定する必要がある場合は使用してよい。

## 変数名を付ける場合

一時変数を作る場合は、可能であれば対象の意味を表す名前にする。

```csharp
var rows = tested.ToArray();
var missingColumns = mapper.FindMissingColumns();
```

単に処理結果であることしか示さない名前より、対象の役割を示す名前を優先する。

テスト対象となるオブジェクトを変数で受ける必要があり、型名や種類名をそのまま変数名にするだけなら `tested` とする。

```csharp
var tested = assembly.GeneratedInstance<Table>("SalesDetailTable");

tested.Name.Should().Be("sales_detail");
tested.Rows.Should().NotBeEmpty();
```

ただし、名前を考えるために不自然な抽象化を追加しない。

英語の変数名を増やすこと自体を可読性向上とはみなさない。日本語話者にとって、式そのものより英語名の方が必ず読みやすいとは限らないためである。

## テスト対象の共有

テストクラス全体の前提になる対象や、多くのテストで同じ初期化が必要な対象は、`readonly` フィールドとしてテストクラスのコンストラクターで初期化する。

```csharp
public sealed class Tableのテスト : IDisposable
{
    readonly Workbook book;
    readonly Table tested;

    public Tableのテスト()
    {
        book = Workbook.Open(@"TestData\Excelテーブル.xlsx");
        tested = book.Tables["Table1"];
    }
}
```

一つのテストでしか使用しない値は、原則としてそのテストメソッド内に置く。

共通化のためだけに定数、フィールド、ヘルパーメソッドを増やさない。

## 期待値の扱い

期待値を一度しか使用せず、長さが許容範囲であればAssert内へ直接書く。

```csharp
tested.Should().BeEquivalentTo(
    [
        new TestRow
        {
            IntegerValue = 1,
            TextValue = "a"
        },
        new TestRow
        {
            IntegerValue = 2,
            TextValue = "b"
        }
    ],
    options => options.WithStrictOrdering());
```

次の場合だけ、期待値を変数へ分離する。

* 複数のAssertで使用する
* 期待値が非常に長く、テストの観測対象が見えにくくなる
* 期待値自体が独立した概念を表す
* 実際値と期待値を別々にデバッグする必要がある

単に `expected` という名前を付けるだけでは、原則として分離する理由にならない。

## 既存テストへの適用

既存テストを整理するときは、次の変更を行う。

* 明示型のローカル変数を、必要がなければ `var` へ変更する
* 一度しか使わない `actual`、`result` などをインライン化する
* 同じ列挙を複数回行っている場合は、一度 `ToArray()` などで受ける
* 同じ例外の複数プロパティを検証している場合は、例外オブジェクトではなく例外アサーションを変数へ受ける
* テストクラス全体の前提になる対象は `readonly` フィールドへ置く
* 一つのテストでしか使わない値はローカルへ戻す
* 変数削減によって式が過度に長くなる場合は無理にインライン化しない
* 振る舞いやテストの意味は変更しない
* テスト名、Skip状態、Skip理由は変更しない
* リファクタリングによってテストの列挙回数や評価タイミングを意図せず変更しない

## 基本原則

一時変数は、英語名を追加するためではなく、次の効果がある場合に使用する。

* 評価回数や列挙回数を固定する
* 同じ検証対象を共有する
* 処理や概念の境界を示す
* デバッグ可能性を高める

これらの効果がない場合は、式を直接読める形を優先する。

ただし、インライン化を目的化しない。

短さよりも、テストがどの振る舞いを検証しているかを明確に読めることを優先する。
