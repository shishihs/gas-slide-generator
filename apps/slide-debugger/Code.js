/**
 * 環境変数の設定
 * この変数を書き換えて、setup() 関数を一度だけ実行してください。
 */
const CONFIG = {
    PRESENTATION_ID: 'ここにスライドIDを入れる',
    FOLDER_ID: 'ここにフォルダIDを入れる' // 省略可
};

/**
 * 初期設定用関数
 * 実行するとスクリプトプロパティに値が保存されます。
 * 以降はソースコードを書き換えなくても、保存された値が使われます。
 */
function setup() {
    const props = PropertiesService.getScriptProperties();

    if (CONFIG.PRESENTATION_ID === 'ここにスライドIDを入れる') {
        Logger.log('エラー: コード上部の CONFIG.PRESENTATION_ID を書き換えてから setup を実行してください。');
        return;
    }

    const properties = {
        'PRESENTATION_ID': CONFIG.PRESENTATION_ID
    };

    if (CONFIG.FOLDER_ID && CONFIG.FOLDER_ID !== 'ここにフォルダIDを入れる') {
        properties['FOLDER_ID'] = CONFIG.FOLDER_ID;
    }

    props.setProperties(properties);
    Logger.log('スクリプトプロパティにより設定が保存されました。exportSlidesToPng を実行できます。');
    Logger.log('設定値:', properties);
}


function exportSlidesToPng() {
    const props = PropertiesService.getScriptProperties();
    const presentationId = props.getProperty('PRESENTATION_ID');
    const folderId = props.getProperty('FOLDER_ID');

    if (!presentationId) {
        Logger.log('エラー: スクリプトプロパティが設定されていません。コード上部のIDを書き換えて setup() 関数を一度実行してください。');
        return;
    }

    const presentation = SlidesApp.openById(presentationId);
    const slides = presentation.getSlides();

    let folder;
    if (!folderId) {
        folder = DriveApp.getRootFolder();
        Logger.log('通知: フォルダIDが設定されていないため、マイドライブ直下に保存します。');
    } else {
        try {
            folder = DriveApp.getFolderById(folderId);
        } catch (e) {
            Logger.log(`エラー: フォルダID (${folderId}) が正しくありません。マイドライブを使用します。`);
            folder = DriveApp.getRootFolder();
        }
    }

    const listName = presentation.getName();
    Logger.log(`プレゼンテーション: ${listName} のエクスポートを開始します`);

    slides.forEach((slide, index) => {
        // 1. メタデータ保持用の透明なシェイプを探す
        let slideType = 'unknown';
        const shapes = slide.getShapes();
        for (const shape of shapes) {
            if (shape.getTitle() === 'SLIDE_METADATA') {
                slideType = shape.getDescription();
                break;
            }
        }

        // 2. なければレイアウト名などから推測
        if (slideType === 'unknown') {
            const layout = slide.getLayout();
            slideType = layout.getLayoutName();
            // レイアウト名も取得できない場合のフォールバック
            if (!slideType) {
                slideType = 'slide';
            }
        }

        // ファイル名用にサニタイズ (英数字とアンダースコア以外は除去)
        slideType = slideType.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase();

        Logger.log(`Slide ${index + 1}: Type=${slideType}`);

        // サムネイル取得と保存
        // サムネイル取得と保存
        try {
            // Advanced Slides Service を使用
            // 1600x900 などの高画質も指定可能ですが、デフォルトで十分な大きさです
            const thumbnailData = Slides.Presentations.Pages.getThumbnail(presentationId, slide.getObjectId(), {
                'thumbnailProperties.mimeType': 'PNG',
                'thumbnailProperties.thumbnailSize': 'LARGE'
            });
            const thumbnailUrl = thumbnailData.contentUrl;

            const response = UrlFetchApp.fetch(thumbnailUrl);
            // 保存形式: {ページ数}_{slide type}.png
            // ページ数は 1-based index
            const fileName = `${index + 1}_${slideType}.png`;
            const blob = response.getBlob().setName(fileName);
            folder.createFile(blob);
        } catch (e) {
            Logger.log(`エラー: Slide ${index + 1} の保存に失敗しました: ${e}`);
        }
    });

    Logger.log(`${slides.length}枚のスライドを保存しました。フォルダ: ${folder.getName()}`);
}
