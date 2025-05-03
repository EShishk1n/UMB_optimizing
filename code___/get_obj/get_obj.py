from code___.get_obj.get_nng_obj import get_nng_obj
from code___.get_obj.get_orn_obj import get_orn_obj
from code___.get_obj.get_smng_obj import get_smng_obj
from code___.get_obj.get_vng_obj import get_vng_obj


def get_obj() -> list:
    obj = get_orn_obj() + get_nng_obj() + get_smng_obj() + get_vng_obj()

    return obj
