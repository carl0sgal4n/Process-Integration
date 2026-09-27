"""Pruebas de invariantes físicos y casos límite. Ejecutar: python -m unittest -v.

Los objetivos aleatorios se contrastan con un balance acumulado independiente,
sin repetir la cascada del programa. Se comprueban todos los equipos y servicios.
"""
import unittest
import numpy as np
from pinch import Stream, EXAMPLE, analyse, validate, lmtd, targets


def independent_heating(streams, dt):
    # A cada nivel T*, lo que falta para cubrir la demanda por encima de ese
    # nivel debe provenir de una utilidad caliente. El máximo fija el mínimo.
    nodes=[v+(-dt/2 if s.hot else dt/2) for s in streams for v in (s.low,s.high)]
    deficits=[]
    for t in nodes:
        h=c=0.
        for s in streams:
            shifted_t=t+(dt/2 if s.hot else -dt/2)
            q=s.cp*max(0.,s.high-max(s.low,shifted_t))
            if s.hot: h+=q
            else: c+=q
        deficits.append(c-h)
    return max(0.,*deficits)


class TestPinch(unittest.TestCase):
    def check_physics(self,r):
        tol=max(1,r['qhot'],r['qcold'])*1e-7
        self.assertAlmostEqual(r['qh'],independent_heating(r['streams'],r['dtmin']),delta=tol)
        self.assertAlmostEqual(r['qhot']+r['qh'],r['qcold']+r['qc'],delta=tol)
        self.assertLessEqual(max(abs(x) for x in r['balance_errors']),tol)
        self.assertGreaterEqual(r['residual'].min(),-tol)
        self.assertAlmostEqual(r['matrix'].sum(),r['recovered'],delta=tol)
        self.assertAlmostEqual(r['cooling'].sum(),r['qc'],delta=tol)
        self.assertAlmostEqual(r['heating'].sum(),r['qh'],delta=tol)
        for e in r['exchangers']:
            self.assertGreaterEqual(min(e['DT_extremo_1'],e['DT_extremo_2']),r['dtmin']-1e-6)
            self.assertAlmostEqual(e['Q_kJh'],e['CP_rama_H']*(e['Th_in']-e['Th_out']),delta=tol)
            self.assertAlmostEqual(e['Q_kJh'],e['CP_rama_C']*(e['Tc_out']-e['Tc_in']),delta=tol)
            for p in r['pinches']:
                # Ningún equipo debe tomar calor caliente por debajo del pinch
                # para una corriente fría por encima, ni cruzar el pinch.
                self.assertFalse(e['Th_in']>p['T_hot']+1e-6 and e['Th_out']<p['T_hot']-1e-6)
                self.assertFalse(e['Tc_in']<p['T_cold']-1e-6 and e['Tc_out']>p['T_cold']+1e-6)

    def test_pdf_reference(self):
        r=analyse(EXAMPLE,10)
        self.assertEqual((r['qh'],r['qc'],r['recovered']),(70000,60000,470000))
        self.assertEqual(r['pinches'][0]['T_hot'],140)
        self.assertEqual(r['pinches'][0]['T_cold'],130)
        np.testing.assert_allclose(r['matrix'][:2,2:],[[32000,98000],[148000,192000]])
        self.check_physics(r)

    def test_changed_dt(self):
        r=analyse(EXAMPLE,20)
        self.assertEqual((r['qh'],r['qc']),(120000,110000))
        self.check_physics(r)

    def test_only_hot_only_cold(self):
        for streams in [[Stream('H',200,200,50)],[Stream('C',300,20,150)]]:
            r=analyse(streams,10)
            self.assertAlmostEqual(r['recovered'],0)
            self.check_physics(r)

    def test_no_thermal_overlap(self):
        r=analyse([Stream('H',1000,60,40),Stream('C',1000,100,150)],10)
        self.assertEqual((r['qh'],r['qc'],r['recovered']),(50000,20000,0))
        self.check_physics(r)

    def test_pinch_zone_and_zero_dt(self):
        for dt in [0,10]:
            r=analyse([Stream('H',1000,200,100),Stream('C',1000,100-dt,200-dt)],dt,1800)
            self.assertAlmostEqual(r['qh']+r['qc'],0)
            self.check_physics(r)
        self.assertEqual(lmtd(10,10),10)
        self.assertEqual(lmtd(0,10),0)

    def test_gaps_and_type_change(self):
        cases=[[Stream('H1',500,300,240),Stream('H2',1500,100,50),Stream('C',900,10,210)],
               [Stream('former_H1',1000,100,250)]+EXAMPLE[1:]]
        for streams in cases:self.check_physics(analyse(streams,10))

    def test_invalid_inputs(self):
        for streams,dt in [([],10),([Stream('x',0,200,50)],10),
            ([Stream('x',1000,50,50)],10),([Stream('x',1000,float('nan'),50)],10),
            ([Stream('x',1000,-300,50)],10),([EXAMPLE[0],EXAMPLE[0]],10),(EXAMPLE,-1)]:
            with self.assertRaises(ValueError):validate(streams,dt)

    def test_randomized_energy_and_temperature_constraints(self):
        rng=np.random.default_rng(7619)
        for _ in range(100):
            streams=[]
            for i in range(int(rng.integers(2,9))):
                low=float(rng.uniform(-50,250)); high=low+float(rng.uniform(5,200))
                tin,tout=(high,low) if rng.random()<.5 else (low,high)
                streams.append(Stream(f'S{i}',float(rng.uniform(50,8000)),tin,tout))
            self.check_physics(analyse(streams,float(rng.uniform(0,120))))


if __name__=='__main__':unittest.main()
