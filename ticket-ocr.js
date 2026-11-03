(function (root) {
  const scriptUrl = 'https://cdn.jsdelivr.net/npm/tesseract.js@6.0.1/dist/tesseract.min.js';
  let libraryPromise;
  async function loadLibrary() {
    if (root.Tesseract) return root.Tesseract;
    if (!libraryPromise) libraryPromise = new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = scriptUrl;
      script.async = true;
      script.onload = () => root.Tesseract ? resolve(root.Tesseract) : reject(new Error('No se pudo cargar el lector de tickets.'));
      script.onerror = () => reject(new Error('No hay conexión para cargar el lector de tickets. Puedes escribir o pegar los productos manualmente.'));
      document.head.append(script);
    });
    return libraryPromise;
  }
  async function readImage(file, onProgress = () => {}) {
    if (!file || !/^image\/(jpeg|png|webp|bmp|tiff)$/i.test(file.type)) throw new Error('Elige una foto JPG, PNG o WebP. Si tu móvil genera HEIC, conviértela a JPG o escribe los productos manualmente.');
    const engine = await loadLibrary();
    const worker = await engine.createWorker('spa', 1, { logger: info => { if (info.status === 'recognizing text') onProgress(Math.round((info.progress || 0) * 100)); } });
    try {
      const result = await worker.recognize(file);
      return String(result?.data?.text || '').trim();
    } finally {
      await worker.terminate();
    }
  }
  root.CocinaTicketOCR = { readImage };
})(window);
