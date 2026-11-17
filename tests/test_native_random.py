"""Check the independent reference vector, repeatability and sampled moments."""
from statistics import mean,pstdev
import unittest
from native_random import MASK,next_word,normal_draws,seed_state


class NativeRandomTests(unittest.TestCase):
    def test_unsigned_32_bit_reference_vector(self):
        # Independently evaluated with uint32_t arithmetic in the C reference.
        state=(1,2,3,4)
        expected=(11520,0,5927040,70819200,2031721883,1637235492,1287239034,3734860849)
        actual=[]
        for _ in expected:
            word,state=next_word(state)
            actual.append(word)
        self.assertEqual(tuple(actual),expected)

    def test_trial_seed_states_are_distinct_and_repeatable(self):
        states=[seed_state(2026,trial) for trial in range(1,1001)]
        self.assertEqual(len(set(states)),1000)
        self.assertTrue(all(any(s) and all(0<=word<=MASK for word in s) for s in states))
        self.assertEqual(list(normal_draws(2026,7,125)),list(normal_draws(2026,7,125)))
        self.assertNotEqual(seed_state(2026,7),seed_state(2027,7))

    def test_normal_moments_and_cross_column_independence(self):
        columns=[[] for _ in range(6)]
        for trial in range(1,1001):
            for row in normal_draws(2026,trial,100):
                for column,value in zip(columns,row):column.append(value)
        for column in columns:
            self.assertLess(abs(mean(column)),.02)
            self.assertLess(abs(pstdev(column)-1),.02)
        for i in range(6):
            for j in range(i):
                a,b=columns[i],columns[j]
                covariance=mean(x*y for x,y in zip(a,b))-mean(a)*mean(b)
                self.assertLess(abs(covariance/(pstdev(a)*pstdev(b))),.02)

    def test_invalid_seed_trial_and_horizon_are_rejected(self):
        for seed in (-1,MASK+1,1.5,True):
            with self.subTest(seed=seed),self.assertRaises(ValueError):seed_state(seed,1)
        for trial in (0,1001,1.5,True):
            with self.subTest(trial=trial),self.assertRaises(ValueError):seed_state(0,trial)
        for years in (-1,126,1.5,True):
            with self.subTest(years=years),self.assertRaises(ValueError):list(normal_draws(0,1,years))
