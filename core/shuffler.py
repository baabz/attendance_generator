import re
import random
from typing import List, Tuple, Dict, Any, Optional
from .models import Student, VenueAllocation


def natural_sort_key(val: Any) -> list:
    """Natural alphanumeric sort key for registration numbers and identifiers."""
    return [int(text) if text.isdigit() else text.upper() for text in re.split(r'(\d+)', str(val).strip())]


def reshuffle_students(
    students: List[Student],
    seed: Optional[Any] = None,
    mode: str = "shuffle"
) -> List[Student]:
    """
    Reshuffle or sort students.
    - mode="shuffle": Random shuffle using provided custom seed (or non-seeded if None)
    - mode="original": Maintain original order
    - mode="reg_no": Sort serially by registration number
    - mode="name": Sort alphabetically by name
    """
    students_copy = list(students)

    if mode == "shuffle":
        if seed is not None and str(seed).strip() != "":
            # Deterministic pseudo-random shuffle based on user's custom seed
            rng = random.Random(str(seed).strip())
            rng.shuffle(students_copy)
        else:
            random.shuffle(students_copy)
    elif mode == "reg_no":
        students_copy.sort(key=lambda s: natural_sort_key(s.reg_no))
    elif mode == "name":
        students_copy.sort(key=lambda s: s.name.upper())

    return students_copy


def allocate_to_venues(
    students: List[Student],
    venues: List[Tuple[str, int]],
    seed: Optional[Any] = None,
    reshuffle: bool = True,
    in_hall_order: str = "reg_no"
) -> Tuple[List[VenueAllocation], Dict[str, Any]]:
    """
    Assign students into designated venues according to the exact capacities provided.

    If reshuffle=True, students are first partitioned randomly across venues based on the seed
    (malpractice prevention separating peer cohorts).
    Then inside each venue, students are arranged according to in_hall_order:
    - 'reg_no': Sorted serially by registration number (best for attendance lookup & roll call)
    - 'original_sn': Sorted by their original class list serial number
    - 'name': Sorted alphabetically by full name
    - 'random': Kept in randomized seating sequence

    Returns:
        Tuple of (allocations, report_summary)
    """
    if reshuffle:
        ordered_students = reshuffle_students(students, seed=seed, mode="shuffle")
    else:
        ordered_students = list(students)

    total_students = len(ordered_students)
    total_capacity = sum(cap for _, cap in venues)

    allocations: List[VenueAllocation] = []
    current_idx = 0

    for venue_name, capacity in venues:
        # Take exactly 'capacity' students for this hall (or remaining if fewer)
        hall_students = ordered_students[current_idx : current_idx + capacity]
        current_idx += len(hall_students)

        # Apply in-hall serial or chosen arrangement
        if in_hall_order == "reg_no":
            hall_students = sorted(hall_students, key=lambda s: natural_sort_key(s.reg_no))
        elif in_hall_order == "original_sn":
            hall_students = sorted(hall_students, key=lambda s: s.original_sn or 0)
        elif in_hall_order == "name":
            hall_students = sorted(hall_students, key=lambda s: s.name.upper())

        allocations.append(VenueAllocation(venue_name=venue_name, capacity=capacity, students=hall_students))

    unallocated_count = max(0, total_students - current_idx)
    unused_capacity = max(0, total_capacity - total_students)

    summary = {
        "total_students": total_students,
        "total_capacity": total_capacity,
        "allocated_students": current_idx,
        "unallocated_students": unallocated_count,
        "unused_capacity": unused_capacity,
        "is_exact_match": total_students == total_capacity,
        "seed_used": str(seed) if seed else "None",
        "in_hall_order": in_hall_order
    }

    return allocations, summary

