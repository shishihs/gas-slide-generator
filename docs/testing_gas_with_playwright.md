# Google Apps Script (GAS) アドオンのテスト戦略：Playwright と Jest の活用

本プロジェクトでは、Google Apps Script (GAS) 特有の制約（ローカルで実行できない、HTMLがiframe内で動作するなど）を克服し、高品質なコードを維持するためのテスト環境を構築しました。
本ドキュメントでは、そのテスト戦略と実装詳細についてまとめます。

## 1. テストの全体像

GASアドオンは大きく分けて以下の3つの層で構成されています。それぞれの層に適したテスト手法を採用しています。

| 層 | 役割 | テスト手法 | ツール |
|---|---|---|---|
| **UI層** | サイドバー/ダイアログのHTML/JS | ブラウザ操作とGAS連携のモックテスト | **Playwright** |
| **GAS Glue層** | `Code.ts` (APIとUIの繋ぎ込み) | Node.js VMによるサンドボックス実行テスト | **Jest** + **vm** |
| **Logic層** | コアライブラリ (`api.ts`など) | 純粋なTypeScript/JavaScript単体テスト | **Jest** |

---

## 2. UI層のテスト (Playwright)

Google Apps ScriptのUI（`HtmlService`）は、実環境では複雑なiframe構造内で動作し、`google.script.run` を通じてサーバー側と通信します。
これをE2Eテストするために、**GAS環境を完全にモック化**する戦略をとりました。

### 戦略: `google.script.run` のモック注入
Playwrightの `page.addScriptTag` やHTMLへの直接注入を利用し、`google` グローバルオブジェクトを模倣します。

```typescript
// テストコード例 (Playwright)
const mockScript = `
  window.google = {
    script: {
      run: {
        withSuccessHandler: function(cb) { this._success = cb; return this; },
        withFailureHandler: function(cb) { this._failure = cb; return this; },
        // サーバー関数のモック
        convertDocumentToJson: function() {
           /* ... 非同期で成功コールバックを呼ぶ ... */ 
        }
      }
    }
  };
`;
// 実際のHTMLにモックを注入してロード
await page.setContent(htmlContent.replace('<head>', '<head>' + mockScript));
```

これにより、GASのバックエンドをデプロイすることなく、ローカル高速にUIの挙動（ローディング表示、エラーハンドリング、画面遷移）を検証できます。

---

## 3. GAS Glue層のテスト (Jest + VM)

トップレベルの関数（例: `onOpen`, `generateSlides`）を含む `Code.ts` は、Apps Scriptランタイムでしか動きません。通常はテスト困難ですが、Node.jsの仮想マシン (`vm`) モジュールを使うことで擬似実行環境を作りました。

### 戦略: VMサンドボックス実行
1. `PropertiesService`, `DriveApp` などのGAS固有サービスのモックを作成。
2. `vm.createContext` でグローバルコンテキストを作成。
3. `Code.ts` (TypeScript) をトランスパイルしてJS化。
4. `vm.runInContext` でコードをロードし、関数をテスト実行。

```typescript
// テストコード例 (Jest)
import * as vm from 'vm';
import * as ts from 'typescript';

// GASサービスのモック
const context = {
  PropertiesService: { /* ... */ },
  DriveApp: { /* ... */ },
  console: console
};

// コードの読み込みと実行
const jsCode = ts.transpileModule(fs.readFileSync('src/Code.ts', 'utf8'), {}).outputText;
vm.createContext(context);
vm.runInContext(jsCode, context);

// 関数テスト
test('generateSlides は正しくライブラリを呼び出す', () => {
  context.generateSlides(mockData);
  expect(context.SlideGeneratorApi.generateSlides).toHaveBeenCalled();
});
```

この手法により、「ライブラリの呼び出しミス」や「プロパティ取得ミス」といった単純だが致命的なバグをデプロイ前に検知できます。

---

## 4. Logic層のテスト (Jest)

純粋なTypeScriptで書かれたライブラリ部分は、標準的なJestのユニットテストでカバーします。
依存関係（`ISlideRepository` など）をインターフェース化し、DI（依存性の注入）を行うことで、Google APIに依存しないテストが可能になっています。

---

## 今後の推奨フロー

開発時は以下のコマンドで各層のテストを実行し、品質を担保することを推奨します。

1. **ロジック変更時**:
   `cd apps/slide-generator-api && npm test`
2. **UI/連携部分変更時**:
   `cd apps/doc-to-slide-addon && npx playwright test`
   `cd apps/doc-to-slide-addon && npx jest` (Glueコードテスト)
