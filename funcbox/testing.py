import time
import io
import contextlib


def evaluate_test_cases(function, test_cases: list[dict], verbose: bool = True, time_unit: str = 'ms'):
    if not callable(function):
        raise ValueError(f"The function '{function}' is not callable.")

    time_unit = 'μs' if time_unit == 'us' else time_unit

    time_multipliers = {
        's': 1,
        'ms': 1e3,
        'μs': 1e6,
        'ns': 1e9
    }

    if time_unit not in time_multipliers.keys():
        raise ValueError(f"The time unit '{time_unit}' is not supported by this function.")

    pass_count = 0
    total_cases = len(test_cases)
    time_multiplier = time_multipliers.get(time_unit, 1e3)

    results = []

    for index, test_case in enumerate(test_cases):
        captured_output = test_result = test_passed = exception = None

        # The function's printing statements are being redirected to this StringIO variable: output_buffer
        output_buffer = io.StringIO()
        with contextlib.redirect_stdout(output_buffer):
            start_time = time.time()
            try:
                test_result = function(**test_case['input'])
            except Exception as e:
                exception = f"{type(e).__name__}: {e}"
            end_time = time.time()

        if not exception:
            captured_output = output_buffer.getvalue()
            test_passed = test_result == test_case['output']
            pass_count += 1 if test_passed else 0

        result = {
            'index': index + 1,
            'input': test_case['input'],
            'captured_output': captured_output,
            'exception': exception,
            'expected': test_case['output'],
            'actual': test_result,
            'passed': test_passed,
            'time': (end_time - start_time) * time_multiplier
        }

        results.append(result)

        if verbose:
            print(f"\033[1mTEST CASE #{index + 1}\033[0m")
            print("Input:", result['input'], sep='\n', end='\n\n')
            print("Captured Output:", f"\033[36m\033[3m{result['captured_output'].rstrip()}\033[0m"
                  if result['captured_output'] else 'NONE', sep='\n', end='\n\n')

            print("Expected Output:", result['expected'], sep='\n', end='\n\n')
            print("Actual Output:", result['actual'], sep='\n', end='\n\n')

            if exception:
                print("Exception:", f"\033[31m{exception}\033[0m", sep='\n', end='\n\n')

            print("Execution Time:",  f"{result['time']:.3f} {time_unit}", sep='\n', end='\n\n')
            print("Test Result:", "\033[32mPASSED\033[0m" if result['passed'] else "\033[31mFAILED\033[0m", sep='\n',
                  end='\n\n\n\n')

    print("\033[1mSUMMARY:\033[0m")
    print("Total: ", total_cases, ",\t\033[32mPASSED: \033[0m", pass_count,
          ",\t\033[31mFAILED: \033[0m", total_cases - pass_count, sep='')

    if not verbose:
        print("\n\033[1mDETAILS:\033[0m\n")
        for result in results:
            print(f"TEST CASE #{result['index']}")
            print(f"Execution Time: {result['time']:.3f} {time_unit}")
            print("Test Result:", "\033[32mPASSED\033[0m" if result['passed'] else "\033[31mFAILED\033[0m")
            print()


if __name__ == '__main__':
    def sample_function(a, b):
        l = [i for i in range(1000000)]
        print(f"The sum of {a} and {b} is {int(a) + int(b)}")
        return a + b

    test_cases = [
        {'input': {'a': '5', 'b': 3}, 'output': 8},
        {'input': {'a': 1, 'b': -1}, 'output': 0},
        {'input': {'a': 10, 'b': 5}, 'output': 15},
    ]

    evaluate_test_cases(sample_function, test_cases, verbose=False, time_unit='us')
