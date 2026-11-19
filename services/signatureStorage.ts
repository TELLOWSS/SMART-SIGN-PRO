import { SignatureFile } from '../types';

const DB_NAME = 'SmartSignProVault';
const DB_VERSION = 1;
const STORE_NAME = 'signatures';

interface StoredSignatureRecord {
  key: string; // `${name}::${variant}`
  name: string;
  variant: string;
  blob: Blob;
  width: number;
  height: number;
  savedAt: number;
}

const openVaultDb = (): Promise<IDBDatabase | null> => {
  if (typeof indexedDB === 'undefined') return Promise.resolve(null);

  return new Promise((resolve) => {
    try {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          db.createObjectStore(STORE_NAME, { keyPath: 'key' });
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => {
        console.warn('[signatureStorage] IndexedDB 열기 실패:', req.error);
        resolve(null);
      };
    } catch (err) {
      console.warn('[signatureStorage] IndexedDB 미지원:', err);
      resolve(null);
    }
  });
};

/**
 * 업로드/직접 작성한 서명 파일을 브라우저 로컬 IndexedDB 보관함에 저장
 * - 100% 사용자 기기 내부에만 저장(외부 서버 전송 없음)
 */
export const saveSignaturesToVault = async (
  entries: { name: string; variant: string; blob: Blob; width: number; height: number }[]
): Promise<void> => {
  if (!entries.length) return;
  const db = await openVaultDb();
  if (!db) return;

  return new Promise((resolve) => {
    try {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      const store = tx.objectStore(STORE_NAME);
      const now = Date.now();
      for (const item of entries) {
        const record: StoredSignatureRecord = {
          key: `${item.name}::${item.variant}`,
          name: item.name,
          variant: item.variant,
          blob: item.blob,
          width: item.width,
          height: item.height,
          savedAt: now,
        };
        store.put(record);
      }
      tx.oncomplete = () => {
        db.close();
        resolve();
      };
      tx.onerror = () => {
        db.close();
        resolve();
      };
    } catch {
      db.close();
      resolve();
    }
  });
};

/**
 * 브라우저 로컬 서명 보관함에서 저장된 모든 서명을 불러와 SignatureFile Map으로 복원
 */
export const loadSignaturesFromVault = async (
  onTrackUrl?: (url: string) => void
): Promise<Map<string, SignatureFile[]>> => {
  const result = new Map<string, SignatureFile[]>();
  const db = await openVaultDb();
  if (!db) return result;

  return new Promise((resolve) => {
    try {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const store = tx.objectStore(STORE_NAME);
      const req = store.getAll();

      req.onsuccess = () => {
        const records = (req.result || []) as StoredSignatureRecord[];
        for (const rec of records) {
          if (!rec.blob || !rec.name || !rec.variant) continue;
          const previewUrl = URL.createObjectURL(rec.blob);
          onTrackUrl?.(previewUrl);
          const list = result.get(rec.name) || [];
          if (!list.some(s => s.variant === rec.variant)) {
            list.push({
              name: rec.name,
              variant: rec.variant,
              previewUrl,
              width: rec.width || 200,
              height: rec.height || 100,
            });
            result.set(rec.name, list);
          }
        }
        db.close();
        resolve(result);
      };

      req.onerror = () => {
        db.close();
        resolve(result);
      };
    } catch {
      db.close();
      resolve(result);
    }
  });
};

/**
 * 브라우저 로컬 서명 보관함 전체 비우기
 */
export const clearSignatureVault = async (): Promise<void> => {
  const db = await openVaultDb();
  if (!db) return;

  return new Promise((resolve) => {
    try {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).clear();
      tx.oncomplete = () => {
        db.close();
        resolve();
      };
      tx.onerror = () => {
        db.close();
        resolve();
      };
    } catch {
      db.close();
      resolve();
    }
  });
};

