from pyxll import error_handler as standard_error_handler
from pyxll import ErrorContext, xlcAlert


def error_handler(context, exc_type, exc_value, exc_traceback):
    if context.error_type == ErrorContext.Type.MACRO:
        msg = str(exc_value)
        if len(msg):
            xlcAlert(msg)
        return standard_error_handler(context, exc_type, exc_value, exc_traceback)
    return standard_error_handler(context, exc_type, exc_value, exc_traceback)