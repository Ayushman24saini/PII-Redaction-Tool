import re

from pii_redactor import (
    is_generated_value,
    luhn_valid,
    valid_phone,
    looks_like_person_name,
)


def test_luhn_valid_credit_card():
    assert luhn_valid("4111-1111-1111-1111")


def test_luhn_invalid_credit_card():
    assert not luhn_valid("4111-1111-1111-1112")


def test_generated_person_placeholder():
    assert is_generated_value("Person 1")
    assert is_generated_value("Person 44")


def test_generated_email_placeholder():
    assert is_generated_value("person1@example.com")


def test_generated_phone_placeholder():
    assert is_generated_value("+91 90000 00001")


def test_generated_organization_placeholder():
    assert is_generated_value("Example Organization 1")


def test_generated_address_placeholder():
    assert is_generated_value("Example Address 1")


def test_real_person_name():
    assert looks_like_person_name("Rakhi Girija Shetty")


def test_real_person_name_with_middle_initial():
    assert looks_like_person_name("Narayana B. Shetty")


def test_invalid_person_phrase():
    assert not looks_like_person_name("Bid Amount")


def test_valid_indian_phone():
    assert valid_phone("+91 9876543210")


def test_invalid_phone():
    assert not valid_phone("12345")


def test_generated_values_are_not_person_names():
    assert not looks_like_person_name("Person 1")