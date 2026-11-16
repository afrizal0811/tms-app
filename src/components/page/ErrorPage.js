// File: src/components/ErrorPage.js
import { removeLocalStorage } from '@/lib/localStorageHandler';
import SelectionLayout from './SelectionLayout';
import { useLanguage } from '@/context/LanguageContext';

export default function ErrorPage() {
  const { t } = useLanguage();
  return (
    <SelectionLayout>
      <div className="flex flex-col items-center justify-center p-8 max-w-lg mx-auto text-center animate-in fade-in zoom-in-95 duration-300">
        <div className="relative mb-6">
          <div className="absolute -inset-4 bg-red-100 rounded-full animate-pulse opacity-50"></div>
          <div className="bg-white p-4 rounded-full shadow-sm border border-red-100 relative">
            <svg
              xmlns="http://www.w3.org/2000/svg"
              className="h-12 w-12 text-red-500"
              fill="none"
              viewBox="0 0 24 24"
              stroke="currentColor"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={1.5}
                d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"
              />
            </svg>
          </div>
          <div className="absolute -right-1 -top-1 bg-red-600 text-white rounded-full p-1 shadow-md">
            <svg
              xmlns="http://www.w3.org/2000/svg"
              className="h-3 w-3"
              viewBox="0 0 20 20"
              fill="currentColor"
            >
              <path
                fillRule="evenodd"
                d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z"
                clipRule="evenodd"
              />
            </svg>
          </div>
        </div>

        <h2 className="text-2xl font-bold text-slate-800 mb-2">{t('error.title')}</h2>
        <p className="text-slate-500 mb-6 leading-relaxed">{t('error.description')}</p>
        <div className="flex flex-col sm:flex-row gap-3 w-full sm:w-auto">
          <button
            onClick={() => {
              removeLocalStorage('data');
              window.location.reload();
            }}
            className="flex items-center justify-center gap-2 px-6 py-2.5 bg-white border border-slate-300 text-slate-600 hover:bg-slate-50 hover:text-red-600 hover:border-red-200 font-medium rounded-lg transition-all text-sm group cursor-pointer"
          >
            <svg
              xmlns="http://www.w3.org/2000/svg"
              fill="none"
              viewBox="0 0 24 24"
              strokeWidth={2}
              stroke="currentColor"
              className="w-4 h-4"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                d="M16.023 9.348h4.992v-.001M2.985 19.644v-4.992m0 0h4.992m-4.993 0l3.181 3.183a8.25 8.25 0 0013.803-3.7M4.031 9.865a8.25 8.25 0 0113.803-3.7l3.181 3.182m0-4.991v4.99"
              />
            </svg>
            {t('common.button.btn_reload')}
          </button>
        </div>
      </div>
    </SelectionLayout>
  );
}
