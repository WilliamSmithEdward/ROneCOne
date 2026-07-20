from __future__ import annotations

import ctypes
import struct
from dataclasses import dataclass, field


MEM_COMMIT = 0x1000
MEM_RESERVE = 0x2000
MEM_RELEASE = 0x8000
PAGE_READWRITE = 0x04
PAGE_EXECUTE_READ = 0x20

STATUS_COMPLETE = 2
STATUS_FAULTED = 3
STATUS_CANCELED = 4
OP_PUSH = 0
OP_ADD = 1
OP_SUBTRACT = 2
OP_MULTIPLY = 3
OP_DIVIDE = 4
OP_NEGATE = 6
OP_EQUAL = 8
OP_NOT_EQUAL = 9
OP_LESS_THAN = 10
OP_LESS_THAN_OR_EQUAL = 11
OP_GREATER_THAN = 12
OP_GREATER_THAN_OR_EQUAL = 13
OP_AND_ALSO = 14
OP_OR_ELSE = 15
OP_NOT = 16
OP_JUMP_IF_FALSE = 100
OP_JUMP_IF_TRUE = 101
OP_TO_BOOLEAN = 102
OP_LOAD_INPUT = 103
BATCH_MAP = 0
BATCH_SUM = 1
BATCH_MIN = 2
BATCH_MAX = 3


@dataclass
class Assembler:
    code: bytearray = field(default_factory=bytearray)
    labels: dict[str, int] = field(default_factory=dict)
    fixups: list[tuple[int, str]] = field(default_factory=list)

    def emit(self, value: str | bytes) -> None:
        if isinstance(value, str):
            self.code.extend(bytes.fromhex(value))
        else:
            self.code.extend(value)

    def label(self, name: str) -> None:
        if name in self.labels:
            raise ValueError(f"Duplicate label: {name}")
        self.labels[name] = len(self.code)

    def jump(self, opcode: str, label: str) -> None:
        self.emit(opcode)
        self.fixups.append((len(self.code), label))
        self.emit(b"\x00\x00\x00\x00")

    def finish(self) -> bytes:
        for offset, label in self.fixups:
            if label not in self.labels:
                raise ValueError(f"Unknown label: {label}")
            displacement = self.labels[label] - (offset + 4)
            self.code[offset : offset + 4] = struct.pack("<i", displacement)
        return bytes(self.code)


def build_kernel() -> bytes:
    """Build the x64 Windows thread-pool callback embedded by ROneCOne.cls."""
    asm = Assembler()

    # Prologue and worker-thread identity. RDX is the thread-pool context.
    asm.emit("53")  # push rbx
    asm.emit("48 83 EC 20")  # sub rsp, 32 (Windows x64 shadow space)
    asm.emit("48 89 D3")  # mov rbx, rdx
    asm.emit("C7 03 01 00 00 00")  # status = running
    asm.emit("48 8B 43 28")  # rax = context->GetCurrentThreadId
    asm.emit("FF D0")  # call rax
    asm.emit("89 43 24")  # context->worker_thread_id = eax
    asm.emit("48 83 C4 20")  # add rsp, 32
    asm.label("setup_program")
    asm.emit("4C 8B 43 10")  # r8 = context->bytecode
    asm.emit("44 8B 4B 08")  # r9d = context->instruction_count
    asm.emit("45 31 D2")  # r10d = stack depth

    asm.label("dispatch")
    asm.emit("83 7B 20 00")  # cancellation requested?
    asm.jump("0F 85", "canceled")
    asm.emit("45 85 C9")  # no instructions remain?
    asm.jump("0F 84", "finish")
    asm.emit("41 8B 00")  # eax = instruction opcode
    asm.emit("49 83 C0 10")  # advance to next 16-byte instruction
    asm.emit("41 FF C9")  # decrement instruction count

    dispatch = (
        (OP_PUSH, "push"),
        (OP_ADD, "add"),
        (OP_SUBTRACT, "subtract"),
        (OP_MULTIPLY, "multiply"),
        (OP_DIVIDE, "divide"),
        (OP_NEGATE, "negate"),
        (OP_EQUAL, "equal"),
        (OP_NOT_EQUAL, "not_equal"),
        (OP_LESS_THAN, "less_than"),
        (OP_LESS_THAN_OR_EQUAL, "less_than_or_equal"),
        (OP_GREATER_THAN, "greater_than"),
        (OP_GREATER_THAN_OR_EQUAL, "greater_than_or_equal"),
        (OP_AND_ALSO, "and_also"),
        (OP_OR_ELSE, "or_else"),
        (OP_NOT, "not"),
        (OP_JUMP_IF_FALSE, "jump_if_false"),
        (OP_JUMP_IF_TRUE, "jump_if_true"),
        (OP_TO_BOOLEAN, "to_boolean"),
        (OP_LOAD_INPUT, "load_input"),
    )
    for opcode, label in dispatch:
        asm.emit("83 F8" + opcode.to_bytes(1, "little").hex())
        asm.jump("0F 84", label)
    asm.jump("E9", "invalid_opcode")

    asm.label("push")
    asm.emit("41 83 FA 40")  # stack capacity is 64 doubles
    asm.jump("0F 83", "stack_overflow")
    asm.emit("F2 41 0F 10 40 F8")  # xmm0 = instruction operand
    asm.emit("F2 42 0F 11 44 D3 40")  # stack[depth] = xmm0
    asm.emit("41 FF C2")  # depth += 1
    asm.jump("E9", "dispatch")

    asm.label("load_input")
    asm.emit("41 83 FA 40")  # stack capacity is 64 doubles
    asm.jump("0F 83", "stack_overflow")
    asm.emit("48 8B 43 30")  # rax = context->input
    asm.emit("8B 8B 44 02 00 00")  # ecx = context->input_index
    asm.emit("F2 0F 10 04 C8")  # xmm0 = input[input_index]
    asm.emit("F2 42 0F 11 44 D3 40")  # stack[depth] = xmm0
    asm.emit("41 FF C2")  # depth += 1
    asm.jump("E9", "dispatch")

    def binary_prelude(name: str) -> None:
        asm.label(name)
        asm.emit("41 83 FA 02")
        asm.jump("0F 82", "stack_underflow")
        asm.emit("41 83 EA 02")  # depth -= 2
        asm.emit("F2 42 0F 10 44 D3 40")  # xmm0 = lhs
        asm.emit("F2 42 0F 10 4C D3 48")  # xmm1 = rhs

    def binary_finish() -> None:
        asm.emit("F2 42 0F 11 44 D3 40")  # replace operands with result
        asm.emit("41 FF C2")  # depth += 1
        asm.jump("E9", "dispatch")

    binary_prelude("add")
    asm.emit("F2 0F 58 C1")
    binary_finish()

    binary_prelude("subtract")
    asm.emit("F2 0F 5C C1")
    binary_finish()

    binary_prelude("multiply")
    asm.emit("F2 0F 59 C1")
    binary_finish()

    binary_prelude("divide")
    asm.emit("66 0F 57 D2")  # xmm2 = 0
    asm.emit("66 0F 2E CA")  # rhs == 0?
    asm.jump("0F 84", "division_by_zero")
    asm.emit("F2 0F 5E C1")
    binary_finish()

    asm.label("negate")
    asm.emit("41 83 FA 01")
    asm.jump("0F 82", "stack_underflow")
    asm.emit("41 FF CA")  # depth -= 1
    asm.emit("F2 42 0F 10 44 D3 40")
    asm.emit("66 0F 57 C9")  # xmm1 = 0
    asm.emit("F2 0F 5C C8")  # xmm1 = -xmm0
    asm.emit("66 0F 28 C1")  # xmm0 = xmm1
    binary_finish()

    comparisons = (
        ("equal", "0F 94 C0"),
        ("not_equal", "0F 95 C0"),
        ("less_than", "0F 92 C0"),
        ("less_than_or_equal", "0F 96 C0"),
        ("greater_than", "0F 97 C0"),
        ("greater_than_or_equal", "0F 93 C0"),
    )
    for label, setcc in comparisons:
        binary_prelude(label)
        asm.emit("66 0F 2E C1")
        asm.emit(setcc)
        asm.emit("0F B6 C0")
        asm.emit("F7 D8")  # VBA Boolean True is -1
        asm.emit("F2 0F 2A C0")
        binary_finish()

    for label, operation in (("and_also", "20 D0"), ("or_else", "08 D0")):
        binary_prelude(label)
        asm.emit("66 0F 57 D2")  # xmm2 = 0
        asm.emit("66 0F 2E C2")
        asm.emit("0F 95 C0")  # al = lhs <> 0
        asm.emit("66 0F 2E CA")
        asm.emit("0F 95 C2")  # dl = rhs <> 0
        asm.emit(operation)
        asm.emit("0F B6 C0")
        asm.emit("F7 D8")
        asm.emit("F2 0F 2A C0")
        binary_finish()

    asm.label("not")
    asm.emit("41 83 FA 01")
    asm.jump("0F 82", "stack_underflow")
    asm.emit("41 FF CA")
    asm.emit("F2 42 0F 10 44 D3 40")
    asm.emit("66 0F 57 C9")
    asm.emit("66 0F 2E C1")
    asm.emit("0F 94 C0")  # al = value == 0
    asm.emit("0F B6 C0")
    asm.emit("F7 D8")
    asm.emit("F2 0F 2A C0")
    binary_finish()

    asm.label("to_boolean")
    asm.emit("41 83 FA 01")
    asm.jump("0F 82", "stack_underflow")
    asm.emit("41 FF CA")
    asm.emit("F2 42 0F 10 44 D3 40")
    asm.emit("66 0F 57 C9")
    asm.emit("66 0F 2E C1")
    asm.emit("0F 95 C0")
    asm.emit("0F B6 C0")
    asm.emit("F7 D8")
    asm.emit("F2 0F 2A C0")
    binary_finish()

    def jump_prelude(name: str) -> None:
        asm.label(name)
        asm.emit("41 83 FA 01")
        asm.jump("0F 82", "stack_underflow")
        asm.emit("41 FF CA")  # temporarily pop the tested value
        asm.emit("F2 42 0F 10 44 D3 40")
        asm.emit("66 0F 57 C9")
        asm.emit("66 0F 2E C1")

    jump_prelude("jump_if_false")
    asm.jump("0F 84", "take_false_jump")
    asm.jump("E9", "dispatch")

    jump_prelude("jump_if_true")
    asm.jump("0F 85", "take_true_jump")
    asm.jump("E9", "dispatch")

    asm.label("take_true_jump")
    asm.emit("B8 FF FF FF FF")
    asm.emit("F2 0F 2A C0")
    asm.emit("F2 42 0F 11 44 D3 40")
    asm.jump("E9", "take_jump")

    asm.label("take_false_jump")
    asm.emit("66 0F 57 C0")
    asm.emit("F2 42 0F 11 44 D3 40")

    asm.label("take_jump")
    asm.emit("41 FF C2")  # retain the tested value as the result
    asm.emit("F2 41 0F 10 40 F8")  # xmm0 = instructions to skip
    asm.emit("F2 0F 2C C0")  # eax = trunc(xmm0)
    asm.emit("41 29 C1")  # remaining instruction count -= eax
    asm.emit("48 C1 E0 04")  # byte offset = instruction count * 16
    asm.emit("49 01 C0")  # advance the instruction pointer
    asm.jump("E9", "dispatch")

    asm.label("finish")
    asm.emit("41 83 FA 01")
    asm.jump("0F 85", "invalid_stack")
    asm.emit("F2 0F 10 43 40")  # xmm0 = stack[0]
    asm.emit("83 BB 40 02 00 00 00")  # context->input_count == 0?
    asm.jump("0F 84", "scalar_finish")
    asm.emit("8B 8B 48 02 00 00")  # ecx = context->batch_mode
    asm.emit("85 C9")
    asm.jump("0F 85", "reduce")
    asm.emit("8B 83 44 02 00 00")  # eax = context->input_index
    asm.emit("4C 8B 5B 38")  # r11 = context->output
    asm.emit("F2 41 0F 11 04 C3")  # output[eax] = xmm0
    asm.jump("E9", "advance_batch")

    asm.label("reduce")
    asm.emit("8B 83 44 02 00 00")  # eax = context->input_index
    asm.emit("85 C0")
    asm.jump("0F 84", "store_reduction")
    asm.emit("83 F9 01")
    asm.jump("0F 84", "sum_reduction")
    asm.emit("83 F9 02")
    asm.jump("0F 84", "min_reduction")
    asm.emit("83 F9 03")
    asm.jump("0F 84", "max_reduction")
    asm.jump("E9", "invalid_opcode")

    asm.label("sum_reduction")
    asm.emit("F2 0F 58 43 18")  # current + accumulated result
    asm.jump("E9", "store_reduction")

    asm.label("min_reduction")
    asm.emit("F2 0F 10 4B 18")
    asm.emit("F2 0F 5D C1")
    asm.jump("E9", "store_reduction")

    asm.label("max_reduction")
    asm.emit("F2 0F 10 4B 18")
    asm.emit("F2 0F 5F C1")

    asm.label("store_reduction")
    asm.emit("F2 0F 11 43 18")

    asm.label("advance_batch")
    asm.emit("8B 83 44 02 00 00")  # eax = context->input_index
    asm.emit("FF C0")  # input_index += 1
    asm.emit("89 83 44 02 00 00")
    asm.emit("3B 83 40 02 00 00")  # more batch inputs?
    asm.jump("0F 82", "setup_program")
    asm.emit("C7 03 02 00 00 00")  # status = complete
    asm.jump("E9", "exit")

    asm.label("scalar_finish")
    asm.emit("F2 0F 11 43 18")  # context->result = xmm0
    asm.emit("C7 03 02 00 00 00")  # status = complete
    asm.jump("E9", "exit")

    def error(label: str, code: int) -> None:
        asm.label(label)
        asm.emit("C7 43 04 " + struct.pack("<I", code).hex())
        asm.emit("C7 03 03 00 00 00")  # status = error
        asm.jump("E9", "exit")

    error("invalid_opcode", 1)
    error("stack_underflow", 2)
    error("stack_overflow", 3)
    error("division_by_zero", 4)
    error("invalid_stack", 5)

    asm.label("canceled")
    asm.emit("C7 03 04 00 00 00")

    asm.label("exit")
    asm.emit("5B C3")  # pop rbx; ret
    return asm.finish()


class NativeContext(ctypes.Structure):
    _fields_ = [
        ("status", ctypes.c_long),
        ("error", ctypes.c_long),
        ("instruction_count", ctypes.c_long),
        ("reserved", ctypes.c_long),
        ("bytecode", ctypes.c_void_p),
        ("result", ctypes.c_double),
        ("cancel", ctypes.c_long),
        ("worker_thread_id", ctypes.c_ulong),
        ("get_current_thread_id", ctypes.c_void_p),
        ("input", ctypes.c_void_p),
        ("output", ctypes.c_void_p),
        ("stack", ctypes.c_double * 64),
        ("input_count", ctypes.c_long),
        ("input_index", ctypes.c_long),
        ("batch_mode", ctypes.c_long),
        ("batch_reserved", ctypes.c_long),
    ]


def instruction(opcode: int, operand: float = 0.0) -> bytes:
    return struct.pack("<i4xd", opcode, operand)


def verify_kernel(kernel: bytes) -> None:
    if ctypes.sizeof(ctypes.c_void_p) != 8:
        raise RuntimeError("The native task kernel requires 64-bit Python on Windows.")

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.VirtualAlloc.argtypes = (
        ctypes.c_void_p,
        ctypes.c_size_t,
        ctypes.c_ulong,
        ctypes.c_ulong,
    )
    kernel32.VirtualAlloc.restype = ctypes.c_void_p
    kernel32.VirtualProtect.argtypes = (
        ctypes.c_void_p,
        ctypes.c_size_t,
        ctypes.c_ulong,
        ctypes.POINTER(ctypes.c_ulong),
    )
    kernel32.VirtualProtect.restype = ctypes.c_int
    kernel32.VirtualFree.argtypes = (ctypes.c_void_p, ctypes.c_size_t, ctypes.c_ulong)
    kernel32.VirtualFree.restype = ctypes.c_int
    kernel32.GetCurrentProcess.restype = ctypes.c_void_p
    kernel32.GetCurrentThreadId.restype = ctypes.c_ulong
    kernel32.FlushInstructionCache.argtypes = (
        ctypes.c_void_p,
        ctypes.c_void_p,
        ctypes.c_size_t,
    )
    kernel32.FlushInstructionCache.restype = ctypes.c_int
    kernel32.CreateThreadpoolWork.argtypes = (
        ctypes.c_void_p,
        ctypes.c_void_p,
        ctypes.c_void_p,
    )
    kernel32.CreateThreadpoolWork.restype = ctypes.c_void_p
    kernel32.SubmitThreadpoolWork.argtypes = (ctypes.c_void_p,)
    kernel32.WaitForThreadpoolWorkCallbacks.argtypes = (ctypes.c_void_p, ctypes.c_int)
    kernel32.CloseThreadpoolWork.argtypes = (ctypes.c_void_p,)

    address = kernel32.VirtualAlloc(
        None, len(kernel), MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE
    )
    if not address:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        ctypes.memmove(address, kernel, len(kernel))
        old_protection = ctypes.c_ulong()
        if not kernel32.VirtualProtect(
            address, len(kernel), PAGE_EXECUTE_READ, ctypes.byref(old_protection)
        ):
            raise ctypes.WinError(ctypes.get_last_error())
        if not kernel32.FlushInstructionCache(
            kernel32.GetCurrentProcess(), address, len(kernel)
        ):
            raise ctypes.WinError(ctypes.get_last_error())

        excel_thread_id = kernel32.GetCurrentThreadId()

        def execute(*program: bytes, canceled: bool = False) -> NativeContext:
            bytecode = b"".join(program)
            bytecode_buffer = ctypes.create_string_buffer(bytecode)
            context = NativeContext()
            context.instruction_count = len(program)
            context.bytecode = ctypes.addressof(bytecode_buffer)
            context.cancel = int(canceled)
            context.get_current_thread_id = ctypes.cast(
                kernel32.GetCurrentThreadId, ctypes.c_void_p
            ).value
            work = kernel32.CreateThreadpoolWork(
                address, ctypes.byref(context), None
            )
            if not work:
                raise ctypes.WinError(ctypes.get_last_error())
            try:
                kernel32.SubmitThreadpoolWork(work)
                kernel32.WaitForThreadpoolWorkCallbacks(work, False)
            finally:
                kernel32.CloseThreadpoolWork(work)
            assert context.worker_thread_id != 0
            assert context.worker_thread_id != excel_thread_id
            return context

        cases = (
            ("add", (2.0, 3.0, OP_ADD), 5.0),
            ("subtract", (7.0, 2.0, OP_SUBTRACT), 5.0),
            ("multiply", (6.0, 7.0, OP_MULTIPLY), 42.0),
            ("divide", (10.0, 4.0, OP_DIVIDE), 2.5),
            ("equal", (4.0, 4.0, OP_EQUAL), -1.0),
            ("not equal", (4.0, 5.0, OP_NOT_EQUAL), -1.0),
            ("less than", (4.0, 5.0, OP_LESS_THAN), -1.0),
            ("less or equal", (5.0, 5.0, OP_LESS_THAN_OR_EQUAL), -1.0),
            ("greater than", (6.0, 5.0, OP_GREATER_THAN), -1.0),
            ("greater or equal", (5.0, 5.0, OP_GREATER_THAN_OR_EQUAL), -1.0),
            ("and", (-1.0, 0.0, OP_AND_ALSO), 0.0),
            ("or", (0.0, -1.0, OP_OR_ELSE), -1.0),
        )
        for name, (left, right, opcode), expected in cases:
            context = execute(
                instruction(OP_PUSH, left),
                instruction(OP_PUSH, right),
                instruction(opcode),
            )
            assert context.status == STATUS_COMPLETE, (name, context.status)
            assert context.error == 0, (name, context.error)
            assert context.result == expected, (name, context.result)

        context = execute(instruction(OP_PUSH, 7.0), instruction(OP_NEGATE))
        assert context.status == STATUS_COMPLETE
        assert context.result == -7.0
        context = execute(instruction(OP_PUSH, 0.0), instruction(OP_NOT))
        assert context.status == STATUS_COMPLETE
        assert context.result == -1.0

        context = execute(
            instruction(OP_PUSH, 0.0),
            instruction(OP_JUMP_IF_FALSE, 4.0),
            instruction(OP_PUSH, 1.0),
            instruction(OP_PUSH, 0.0),
            instruction(OP_DIVIDE),
            instruction(OP_TO_BOOLEAN),
        )
        assert context.status == STATUS_COMPLETE
        assert context.result == 0.0
        context = execute(
            instruction(OP_PUSH, 1.0),
            instruction(OP_JUMP_IF_TRUE, 4.0),
            instruction(OP_PUSH, 1.0),
            instruction(OP_PUSH, 0.0),
            instruction(OP_DIVIDE),
            instruction(OP_TO_BOOLEAN),
        )
        assert context.status == STATUS_COMPLETE
        assert context.result == -1.0

        context = execute(
            instruction(OP_PUSH, 1.0),
            instruction(OP_PUSH, 0.0),
            instruction(OP_DIVIDE),
        )
        assert context.status == STATUS_FAULTED
        assert context.error == 4

        context = execute(instruction(OP_PUSH, 1.0), canceled=True)
        assert context.status == STATUS_CANCELED

        inputs = (ctypes.c_double * 4)(-2.0, 0.0, 3.0, 10.0)
        outputs = (ctypes.c_double * 4)()
        bytecode = b"".join(
            (
                instruction(OP_LOAD_INPUT),
                instruction(OP_PUSH, 2.0),
                instruction(OP_MULTIPLY),
            )
        )
        bytecode_buffer = ctypes.create_string_buffer(bytecode)
        context = NativeContext()
        context.instruction_count = 3
        context.bytecode = ctypes.addressof(bytecode_buffer)
        context.get_current_thread_id = ctypes.cast(
            kernel32.GetCurrentThreadId, ctypes.c_void_p
        ).value
        context.input = ctypes.addressof(inputs)
        context.output = ctypes.addressof(outputs)
        context.input_count = len(inputs)
        work = kernel32.CreateThreadpoolWork(address, ctypes.byref(context), None)
        if not work:
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            kernel32.SubmitThreadpoolWork(work)
            kernel32.WaitForThreadpoolWorkCallbacks(work, False)
        finally:
            kernel32.CloseThreadpoolWork(work)
        assert context.status == STATUS_COMPLETE
        assert context.input_index == len(inputs)
        assert list(outputs) == [-4.0, 0.0, 6.0, 20.0]

        def reduce_batch(mode: int) -> NativeContext:
            bytecode = instruction(OP_LOAD_INPUT)
            bytecode_buffer = ctypes.create_string_buffer(bytecode)
            reduced = NativeContext()
            reduced.instruction_count = 1
            reduced.bytecode = ctypes.addressof(bytecode_buffer)
            reduced.get_current_thread_id = ctypes.cast(
                kernel32.GetCurrentThreadId, ctypes.c_void_p
            ).value
            reduced.input = ctypes.addressof(inputs)
            reduced.input_count = len(inputs)
            reduced.batch_mode = mode
            work = kernel32.CreateThreadpoolWork(
                address, ctypes.byref(reduced), None
            )
            if not work:
                raise ctypes.WinError(ctypes.get_last_error())
            try:
                kernel32.SubmitThreadpoolWork(work)
                kernel32.WaitForThreadpoolWorkCallbacks(work, False)
            finally:
                kernel32.CloseThreadpoolWork(work)
            assert reduced.status == STATUS_COMPLETE
            return reduced

        assert reduce_batch(BATCH_SUM).result == 11.0
        assert reduce_batch(BATCH_MIN).result == -2.0
        assert reduce_batch(BATCH_MAX).result == 10.0
    finally:
        kernel32.VirtualFree(address, 0, MEM_RELEASE)


def format_vba_hex(kernel: bytes, width: int = 96) -> str:
    value = kernel.hex().upper()
    return "\n".join(value[index : index + width] for index in range(0, len(value), width))


if __name__ == "__main__":
    built_kernel = build_kernel()
    verify_kernel(built_kernel)
    print(f"Verified native x64 task kernel: {len(built_kernel)} bytes")
    print(format_vba_hex(built_kernel))
