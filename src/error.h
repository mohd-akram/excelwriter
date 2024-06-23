#include <napi.h>
#include <xlsxwriter.h>

inline void checkError(Napi::Env env, lxw_error error) {
  if (error != LXW_NO_ERROR)
    throw Napi::Error::New(env, lxw_strerror(error));
}
