import { toast } from "react-toastify";

export const reviewConfirm = (onConfirm) => {
    toast(
        ({ closeToast }) => (
            <div>
                <p>Have you reviewed all BOM data? Once submitted, the BOM and Raw Materials will be created.</p>

                <div className="flex gap-2 mt-2">

                    <button
                        className="bg-red-500 text-white px-2 py-1 rounded"
                        onClick={() => {
                            onConfirm();
                            closeToast();
                        }}
                    >
                        Yes
                    </button>

                    <button
                        className="bg-gray-500 text-white px-2 py-1 rounded"
                        onClick={closeToast}
                    >
                        No
                    </button>

                </div>
            </div>
        ),
        {
            autoClose: false,
            closeOnClick: false
        }
    );
};