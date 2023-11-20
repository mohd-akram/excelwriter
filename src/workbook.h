#include <napi.h>
#include <xlsxwriter.h>

class Workbook : public Napi::ObjectWrap<Workbook> {
 public:
  static Napi::Object Init(Napi::Env env, Napi::Object exports);
  Workbook(const Napi::CallbackInfo& info);
  ~Workbook();

 private:
  Napi::Value AddChart(const Napi::CallbackInfo& info);
  Napi::Value AddFormat(const Napi::CallbackInfo& info);
  Napi::Value AddWorksheet(const Napi::CallbackInfo& info);
  Napi::Value GetDefaultURLFormat(const Napi::CallbackInfo& info);
  Napi::Value Close(const Napi::CallbackInfo& info);
  Napi::ObjectReference default_url_format;
  lxw_workbook* workbook = nullptr;
  char* output_buffer = nullptr;
  size_t output_buffer_size = 0;
};
