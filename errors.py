# errors.py — updated
from functools import wraps
from flask import g, request
import logging

logger = logging.getLogger("tefteri")


class AppError(Exception):
    def __init__(self, message, status_code=400):
        super().__init__(message)
        self.message = message
        self.status_code = status_code


class NotFoundError(AppError):
    def __init__(self, message="The requested item was not found."):
        super().__init__(message, status_code=404)


class ValidationError(AppError):
    def __init__(self, message="Please check your input and try again."):
        super().__init__(message, status_code=400)


class OwnershipError(AppError):
    def __init__(self, message="You don't have access to that item."):
        super().__init__(message, status_code=403)


class AuthError(AppError):
    """For routes that manually check `session["user_id"]` instead of
    using @login_required (e.g. because they need a custom 401 JSON
    response rather than a redirect to /login)."""
    def __init__(self, message="Please log in to continue."):
        super().__init__(message, status_code=401)


def json_route(view):
    """Decorator: mark a route as JSON-only. Any AppError raised inside
    always returns JSON, regardless of the request's Accept header."""
    @wraps(view)
    def wrapper(*args, **kwargs):
        g.force_json = True
        return view(*args, **kwargs)
    return wrapper


def wants_json():
    if getattr(g, "force_json", False):
        return True
    if request.headers.get("X-Requested-With") == "XMLHttpRequest":
        return True
    if request.is_json:
        return True
    return False