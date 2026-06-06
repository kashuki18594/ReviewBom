import { toast } from "react-toastify";

export const confirmRejectBom = (onConfirm) => {
    toast(
        ({ closeToast }) => (
            <div>
                <p>Are you sure you want to Reject BOM?</p>

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