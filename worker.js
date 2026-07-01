/* worker.js — マルチファイル配信用のWorkerエントリ。
 * ライブラリ・コア・処理本体を読み込む。単一版ではBlobで代替される。 */
importScripts('vendor/jszip.min.js', 'vendor/exceljs.min.js', 'crud-core.js', 'worker-core.js');
