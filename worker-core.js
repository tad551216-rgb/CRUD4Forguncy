/* worker-core.js — バックグラウンド処理本体（self.CrudCore を使用）
 * worker.js から importScripts で読み込まれるか、単一版ではBlobに連結される。 */
(function () {
  'use strict';

  function toArrayBuffer(x) {
    if (x instanceof ArrayBuffer) return x;
    if (x && x.buffer instanceof ArrayBuffer) {
      return x.buffer.slice(x.byteOffset || 0, (x.byteOffset || 0) + x.byteLength);
    }
    return new Uint8Array(x).buffer;
  }

  self.onmessage = async function (e) {
    const msg = e.data || {};
    const post = (o, transfer) => self.postMessage(o, transfer || []);
    try {
      if (!self.CrudCore) throw new Error('処理モジュールの読み込みに失敗しました');

      if (msg.cmd === 'diagnose') {
        const d = await self.CrudCore.diagnose(msg.fgcpBuffer, (done, total) =>
          post({ type: 'progress', ev: { stage: 'parse', done, total } }));
        post({ type: 'diagnose-done', result: d });

      } else if (msg.cmd === 'generate') {
        const r = await self.CrudCore.generate({
          fgcpBuffer: msg.fgcpBuffer,
          fgdocBuffer: msg.fgdocBuffer || null,
          onProgress: (ev) => post({ type: 'progress', ev }),
        });
        const xbuf = toArrayBuffer(await r.workbook.xlsx.writeBuffer());
        post({
          type: 'generate-done',
          stats: r.stats, mermaid: r.mermaid, orphanTables: r.orphanTables, xlsx: xbuf,
        }, [xbuf]);

      } else {
        throw new Error('不明な指示: ' + msg.cmd);
      }
    } catch (err) {
      post({ type: 'error', message: (err && err.message) ? err.message : String(err) });
    }
  };
})();
