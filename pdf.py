from docx import Document
import re

def create_mcq_bank():
    # This is the text you provided
    text_content = """
Since I cannot directly provide a downloadable `.docx` or `.pdf` file in this interface, I have compiled all **215+ questions** (including the previous 200 and the new diagram-based ones) into a structured format below. 

You can **copy and paste** the content directly into Microsoft Word. I have formatted it with clear headings and spacing for easy reading.

***

# Electric Circuit Analysis: MCQ Bank for Core Companies
**Comprehensive Question Bank (Basic to Advanced)**

## **Section 1: Basic Circuit Analysis (Units, Components, Laws)**

1. **Which of the following is not an electrical quantity?**
   A) Voltage
   B) Current
   C) Power
   D) Distance
   **Answer: D**

2. **The S.I. unit of electric current is:**
   A) Volt
   B) Ampere
   C) Ohm
   D) Watt
   **Answer: B**

3. **One Volt is equivalent to:**
   A) 1 Joule per Coulomb
   B) 1 Coulomb per Joule
   C) 1 Watt per Ampere
   D) Both A and C
   **Answer: D**

4. **If a charge of 10 Coulombs flows past a point in 2 seconds, the current is:**
   A) 20 A
   B) 5 A
   C) 0.2 A
   D) 10 A
   **Answer: B**

5. **The resistance of a conductor is directly proportional to:**
   A) Length
   B) Cross-sectional area
   C) Temperature coefficient
   D) Current
   **Answer: A**

6. **Ohm's Law states that:**
   A) V = IR
   B) V = I/R
   C) I = VR
   D) R = VI
   **Answer: A**

7. **Which component opposes the flow of current?**
   A) Inductor
   B) Capacitor
   C) Resistor
   D) Voltage Source
   **Answer: C**

8. **In a series circuit, the total resistance is:**
   A) Less than the smallest resistance
   B) The sum of individual resistances
   C) The product of individual resistances
   D) Constant
   **Answer: B**

9. **In a parallel circuit, the voltage across each branch is:**
   A) Different
   B) The same
   C) Zero
   D) Divided by branch resistance
   **Answer: B**

10. **Kirchhoff's Current Law (KCL) is applicable to:**
    A) Closed loops
    B) Junction points (Nodes)
    C) Magnetic circuits
    D) Open circuits
    **Answer: B**

11. **Kirchhoff's Voltage Law (KVL) is based on the conservation of:**
    A) Charge
    B) Mass
    C) Energy
    D) Momentum
    **Answer: C**

12. **Calculate the total resistance for two resistors of 4Ω and 6Ω connected in series.**
    A) 10Ω
    B) 2.4Ω
    C) 24Ω
    D) 1.5Ω
    **Answer: A**

13. **Calculate the equivalent resistance for two resistors of 4Ω and 6Ω connected in parallel.**
    A) 10Ω
    B) 2.4Ω
    C) 24Ω
    D) 1.5Ω
    **Answer: B**

14. **A passive element is one which:**
    A) Supplies energy
    B) Absorbs energy
    C) Both supplies and absorbs
    D) None of the above
    **Answer: B**

15. **The power dissipated in a resistor is given by:**
    A) $P = VI$
    B) $P = I^2R$
    C) $P = V^2/R$
    D) All of the above
    **Answer: D**

16. **An ideal voltage source has:**
    A) Zero internal resistance
    B) Infinite internal resistance
    C) Finite internal resistance
    D) None of these
    **Answer: A**

17. **An ideal current source has:**
    A) Zero internal resistance
    B) Infinite internal resistance
    C) Finite internal resistance
    D) None of these
    **Answer: B**

18. **A 100W bulb is connected to a 200V supply. The current drawn is:**
    A) 2A
    B) 0.5A
    C) 20A
    D) 5A
    **Answer: B**

19. **Mesh analysis is based on:**
    A) KCL
    B) KVL
    C) Ohm's Law
    D) Faraday's Law
    **Answer: B**

20. **Nodal analysis is based on:**
    A) KCL
    B) KVL
    C) Ohm's Law
    D) Coulomb's Law
    **Answer: A**

21. **A circuit diagram shows a 5A current source in parallel with a 10Ω resistor. Converting this to a voltage source results in:**
    A) 50V in series with 10Ω
    B) 0.5V in series with 10Ω
    C) 50V in parallel with 10Ω
    D) 2V in series with 10Ω
    **Answer: A**

22. **The number of independent loops in a circuit is given by:**
    A) $b - n + 1$
    B) $b + n - 1$
    C) $n - 1$
    D) $b - 1$
    *(Where b = branches, n = nodes)*
    **Answer: A**

23. **If the voltage across a capacitor is constant, the current through it is:**
    A) Infinite
    B) Zero
    C) Maximum
    D) Half the voltage
    **Answer: B**

24. **The current through an inductor cannot change instantaneously because:**
    A) It requires infinite voltage
    B) Power would be infinite
    C) Flux linkage cannot change instantly
    D) All of the above
    **Answer: D**

25. **Which of the following is a bilateral element?**
    A) Diode
    B) Resistor
    C) Transistor
    D) Voltage Source
    **Answer: B**

26. **The V-I relationship for a capacitor is:**
    A) $i = C \frac{di}{dt}$
    B) $i = C \frac{dv}{dt}$
    C) $v = C \frac{di}{dt}$
    D) $v = L \frac{di}{dt}$
    **Answer: B**

27. **Energy stored in a capacitor is:**
    A) $\frac{1}{2} CV^2$
    B) $\frac{1}{2} LI^2$
    C) $CV^2$
    D) $LI^2$
    **Answer: A**

28. **Energy stored in an inductor is:**
    A) $\frac{1}{2} CV^2$
    B) $\frac{1}{2} LI^2$
    C) $CV^2$
    D) $LI^2$
    **Answer: B**

29. **In a series circuit with R and L, the phase angle between voltage and current is:**
    A) 0 degrees
    B) 90 degrees (Current lags)
    C) 90 degrees (Current leads)
    D) Depends on values
    **Answer: B**

30. **A capacitor blocks DC because its reactance is:**
    A) Zero
    B) Infinite
    C) Very low
    D) Negative
    **Answer: B**

---

## **Section 2: Network Reduction and Theorems (DC & AC)**

31. **The Superposition Theorem is applicable to:**
    A) Linear circuits only
    B) Non-linear circuits only
    C) Both linear and non-linear
    D) Passive circuits only
    **Answer: A**

32. **Thevenin’s Theorem is used to reduce a complex circuit to a:**
    A) Voltage source in parallel with a resistor
    B) Current source in series with a resistor
    C) Voltage source in series with a resistor
    D) Current source in parallel with a resistor
    **Answer: C**

33. **Norton’s Theorem reduces a circuit to a:**
    A) Voltage source in series with a resistor
    B) Current source in parallel with a resistor
    C) Voltage source in parallel with a resistor
    D) Current source in series with a resistor
    **Answer: B**

34. **Maximum Power Transfer Theorem states that maximum power is delivered to the load when:**
    A) Load Resistance = Source Resistance
    B) Load Resistance = 0
    C) Load Resistance = Infinity
    D) Load Resistance = Internal Reactance
    **Answer: A**

35. **In AC circuits, for Maximum Power Transfer, the load impedance should be:**
    A) Complex conjugate of source impedance
    B) Equal to source impedance
    C) Double the source impedance
    D) Half the source impedance
    **Answer: A**

36. **A Delta network of 3 equal resistors R is converted to Star. The resistance in each branch of Star is:**
    A) R
    B) 3R
    C) R/3
    D) 2R
    **Answer: C**

37. **A Star network of 3 equal resistors R is converted to Delta. The resistance in each branch of Delta is:**
    A) R
    B) 3R
    C) R/3
    D) 2R
    **Answer: B**

38. **Problem: Three resistors 3Ω, 3Ω, 3Ω are connected in Delta. The equivalent Star resistance is:**
    A) 9Ω
    B) 3Ω
    C) 1Ω
    D) 6Ω
    **Answer: C**

39. **Reciprocity Theorem is applicable to:**
    A) Single source circuits
    B) Multi-source circuits
    C) Passive bilateral networks
    D) Both A and C
    **Answer: D**

40. **The Superposition Theorem is not applicable to:**
    A) Voltage
    B) Current
    C) Power
    D) Resistance
    **Answer: C** (Power is non-linear function of V/I)

41. **To find Thevenin voltage ($V_{th}$):**
    A) Short circuit the load
    B) Open circuit the load
    C) Remove sources
    D) Short circuit voltage sources
    **Answer: B**

42. **To find Thevenin resistance ($R_{th}$) in a circuit with dependent sources:**
    A) Open circuit voltage sources and short circuit current sources
    B) Connect a 1V source at the terminals and find current
    C) Use $V_{th}/I_{sc}$
    D) All of the above methods work
    **Answer: D**

43. **Millman’s Theorem is used to solve circuits having:**
    A) Parallel voltage sources
    B) Series current sources
    C) Only resistors
    D) Complex impedance
    **Answer: A**

44. **Tellegen’s Theorem is based on conservation of:**
    A) Charge
    B) Energy
    C) Momentum
    D) Flux
    **Answer: B**

45. **In a simple AC circuit with a pure inductor, the power consumed is:**
    A) Maximum
    B) Zero
    C) Minimum
    D) Depends on frequency
    **Answer: B**

46. **Impedance of a series RLC circuit at resonance is:**
    A) Maximum
    B) Minimum
    C) Equal to R
    D) Both B and C
    **Answer: D**

47. **Impedance of a parallel RLC circuit at resonance is:**
    A) Minimum
    B) Maximum
    C) Equal to R
    D) Zero
    **Answer: B**

48. **Concept Question: A circuit has $V_{th}=10V$ and $R_{th}=5\Omega$. What is the Norton current?**
    A) 2A
    B) 0.5A
    C) 50A
    D) 10A
    **Answer: A**

49. **Concept Question: If a circuit has a short circuit across the load, the voltage across the load is:**
    A) Infinite
    B) Source Voltage
    C) Zero
    D) Negative
    **Answer: C**

50. **Which theorem assists in replacing a current source in parallel with a resistor by a voltage source in series with a resistor?**
    A) Thevenin's
    B) Norton's
    C) Source Transformation
    D) Superposition
    **Answer: C**

---

## **Section 3: Nodal and Mesh Analysis (DC & AC)**

51. **In Nodal Analysis, the reference node is chosen to have a potential of:**
    A) +1V
    B) -1V
    C) 0V
    D) Infinity
    **Answer: C**

52. **The number of nodal equations required for a circuit with $n$ nodes is:**
    A) $n$
    B) $n-1$
    C) $n+1$
    D) $2n$
    **Answer: B**

53. **If a voltage source is connected between two non-reference nodes, the method to solve it is:**
    A) Supermesh
    B) Supernode
    C) Source Transformation
    D) Both B and C
    **Answer: D**

54. **If a current source lies on the perimeter of a mesh, it simplifies the analysis because:**
    A) The mesh current equals the source current
    B) The voltage drop is zero
    C) It must be removed
    D) It creates a supermesh
    **Answer: A**

55. **Supermesh is formed when two meshes share a:**
    A) Voltage source
    B) Current source
    C) Resistor
    D) Capacitor
    **Answer: B**

56. **Problem (Diagram Description): A circuit has a 10V source in series with 2Ω connected to Node A. Node A is also connected to a 5Ω resistor going to ground. What is the voltage at Node A?**
    A) 10V
    B) 0V
    C) 5V
    D) 7.14V
    **Answer: D** (Using voltage division: $10 \times \frac{5}{2+5}$)

57. **In Nodal Analysis, the current leaving a node through a resistor R is written as (assuming $V_{node} > V_{neighbor}$):**
    A) $(V_{neighbor} - V_{node}) / R$
    B) $(V_{node} - V_{neighbor}) / R$
    C) $V_{node} / R$
    D) $V_{neighbor} / R$
    **Answer: B**

58. **Matrix form of Nodal Analysis equations is given by:**
    A) $[I] = [R][V]$
    B) $[V] = [R][I]$
    C) $[I] = [G][V]$
    D) $[G] = [V][I]$
    *(Where G is conductance)*
    **Answer: C**

59. **Problem: In a 3-node circuit (excluding reference), the conductance matrix diagonal element $G_{22}$ represents:**
    A) Sum of conductances connected to node 1
    B) Sum of conductances connected to node 2
    C) Negative of mutual conductance
    D) Resistance of node 2
    **Answer: B**

60. **For AC Nodal Analysis, the impedances are treated as:**
    A) Real numbers
    B) Complex numbers
    C) Phasors
    D) Both B and C
    **Answer: D**

61. **Problem: Calculate the current through a 2Ω resistor using Mesh Analysis. Mesh 1 current ($I_1$) = 3A (clockwise), Mesh 2 current ($I_2$) = 1A (clockwise). The resistor is common to both meshes.**
    A) 2A
    B) 4A
    C) 3A
    D) 1A
    **Answer: A** ($I_{resistor} = I_1 - I_2$)

62. **Which analysis method is generally preferred for a circuit with fewer nodes than meshes?**
    A) Mesh Analysis
    B) Nodal Analysis
    C) Superposition
    D) Thevenin's
    **Answer: B**

63. **Which analysis method is preferred for a circuit with fewer meshes than nodes?**
    A) Nodal Analysis
    B) Mesh Analysis
    C) Millman's
    D) Source Transformation
    **Answer: B**

64. **In the context of Load Flow studies in power systems (advanced nodal analysis), the bus where voltage magnitude and angle are specified is called:**
    A) PQ Bus
    B) PV Bus
    C) Slack Bus / Reference Bus
    D) Load Bus
    **Answer: C**

65. **In Load Flow analysis, the variable solved for at a "Load Bus" (PQ Bus) is:**
    A) P and Q
    B) V and $\delta$ (Angle)
    C) Q and V
    D) P and V
    **Answer: B**

66. **The term "Admittance Matrix" ($Y_{bus}$) is primarily used in which analysis technique?**
    A) Mesh Analysis
    B) Nodal Analysis
    C) Thevenin Analysis
    D) Transient Analysis
    **Answer: B**

67. **If a dependent source is present in the circuit, Nodal/Mesh analysis requires:**
    A) Ignoring the dependent source
    B) Treating it as an independent source
    C) Writing the controlling equation as an additional constraint
    D) Removing the element
    **Answer: C**

68. **Problem: A voltage source of $10\angle0^\circ$ V is applied to a series R-L circuit ($R=3\Omega, X_L=4\Omega$). The current phasor is:**
    A) $2\angle53.13^\circ$ A
    B) $2\angle-53.13^\circ$ A
    C) $2.5\angle0^\circ$ A
    D) $10\angle53.13^\circ$ A
    **Answer: B** (Impedance $Z = 5\Omega$. $I = V/Z = 2A$. Current lags voltage by $\tan^{-1}(4/3)$).

69. **For a planar circuit, "Mesh" is defined as a loop that contains:**
    A) Any number of inner loops
    B) No inner loops
    C) Only one resistor
    D) A current source
    **Answer: B**

70. **What is the phase difference between the applied voltage and current in a series resonant circuit?**
    A) $90^\circ$
    B) $0^\circ$
    C) $45^\circ$
    D) $180^\circ$
    **Answer: B**

---

## **Section 4: Transient Response Analysis**

71. **The Laplace Transform of a unit step function $u(t)$ is:**
    A) $s$
    B) $1/s$
    C) $1$
    D) $s^2$
    **Answer: B**

72. **The Laplace Transform of $e^{-at}$ is:**
    A) $1/(s+a)$
    B) $1/(s-a)$
    C) $s+a$
    D) $s-a$
    **Answer: A**

73. **The Initial Value Theorem in Laplace Transform is used to find $f(0)$ from $F(s)$ as:**
    A) $\lim_{s \to 0} sF(s)$
    B) $\lim_{s \to \infty} sF(s)$
    C) $\lim_{s \to \infty} F(s)$
    D) $\lim_{s \to 0} F(s)$
    **Answer: B**

74. **The Final Value Theorem is used to find $f(\infty)$ from $F(s)$ as:**
    A) $\lim_{s \to 0} sF(s)$
    B) $\lim_{s \to \infty} sF(s)$
    C) $\lim_{s \to 0} F(s)$
    D) $\lim_{s \to \infty} F(s)$
    **Answer: A**

75. **In an RL circuit, the time constant is given by:**
    A) $R/L$
    B) $L/R$
    C) $RL$
    D) $1/RL$
    **Answer: B**

76. **In an RC circuit, the time constant is given by:**
    A) $RC$
    B) $R/C$
    C) $C/R$
    D) $1/RC$
    **Answer: A**

77. **The transient response in an RLC circuit is underdamped if:**
    A) $\alpha = \omega_0$
    B) $\alpha > \omega_0$
    C) $\alpha < \omega_0$
    D) $\alpha = 0$
    *(Where $\alpha$ is damping factor, $\omega_0$ is resonant frequency)*
    **Answer: C**

78. **The transient response in an RLC circuit is critically damped if:**
    A) $\alpha = \omega_0$
    B) $\alpha > \omega_0$
    C) $\alpha < \omega_0$
    D) $\alpha = \infty$
    **Answer: A**

79. **In a source-free RL circuit, the current decays exponentially with a time constant of:**
    A) $L/R$ seconds
    B) $R/L$ seconds
    C) $1/(L/R)$ seconds
    D) None of the above
    **Answer: A**

80. **The Laplace transform of a derivative $\frac{df}{dt}$ is:**
    A) $sF(s)$
    B) $sF(s) - f(0)$
    C) $F(s)/s$
    D) $F(s) - f(0)$
    **Answer: B**

81. **The inverse Laplace transform of $1/s^2$ is:**
    A) $t$
    B) $t^2$
    C) $1/t$
    D) $u(t)$
    **Answer: A**

82. **Problem: A 10H inductor carries a current of 20mA. How long does it take to interrupt the current if a voltage of 100V is induced?**
    A) 1ms
    B) 2ms
    C) 0.5ms
    D) 20ms
    **Answer: B** ($V = L \frac{di}{dt} \Rightarrow dt = L \frac{di}{V} = 10 \times 0.02 / 100 = 0.002s$)

83. **In a parallel RLC circuit, the voltage across the capacitor at $t=0^+$ is:**
    A) Zero
    B) Maximum
    C) Equal to source voltage
    D) Cannot be determined
    **Answer: C** (Voltage is continuous)

84. **The concept of "Poles" in Laplace transform determines the:**
    A) Stability and nature of response
    B) Amplitude of the signal
    C) Phase of the signal
    D) Frequency of oscillation only
    **Answer: A**

85. **If the poles of a system function lie on the imaginary axis, the response is:**
    A) Decaying
    B) Rising
    C) Sustained oscillation
    D) Unbounded
    **Answer: C**

86. **The impedance of an inductor $L$ in the s-domain is:**
    A) $Ls$
    B) $1/Ls$
    C) $L/s$
    D) $s/L$
    **Answer: A**

87. **The impedance of a capacitor $C$ in the s-domain is:**
    A) $Cs$
    B) $1/Cs$
    C) $C/s$
    D) $s/C$
    **Answer: B**

88. **Problem: An RC circuit has $R=1k\Omega, C=1\mu F$. The time constant is:**
    A) 1s
    B) 1ms
    C) 100s
    D) 0.1s
    **Answer: B** ($\tau = 10^3 \times 10^{-6} = 10^{-3}s$)

89. **Sinusoidal steady-state response implies:**
    A) The transient part has died out
    B) The response is a constant DC value
    C) The circuit is resonant
    D) The source is DC
    **Answer: A**

90. **The "Damping Ratio" ($\zeta$) for a series RLC circuit is:**
    A) $R/2 \sqrt{L/C}$
    B) $\sqrt{L/C} / R$
    C) $R/L$
    D) $1/RC$
    **Answer: A**

---

## **Section 5: Resonance and Coupled Circuits**

91. **Resonance occurs in an RLC circuit when:**
    A) $X_L = X_C$
    B) $X_L > X_C$
    C) $X_L < X_C$
    D) $R = X_L$
    **Answer: A**

92. **The resonant frequency ($f_0$) of a series RLC circuit is:**
    A) $2\pi\sqrt{LC}$
    B) $\frac{1}{2\pi\sqrt{LC}}$
    C) $\frac{1}{2\pi} \sqrt{\frac{L}{C}}$
    D) $\frac{1}{2\pi} \sqrt{\frac{C}{L}}$
    **Answer: B**

93. **At series resonance, the power factor is:**
    A) 0
    B) 0.5
    C) 1
    D) Lagging
    **Answer: C**

94. **In a series resonant circuit, the voltage across the inductor or capacitor can be higher than the source voltage. This phenomenon is called:**
    A) Resonance
    B) Quality Factor
    C) Voltage Magnification
    D) Selectivity
    **Answer: C**

95. **Quality Factor (Q) of a series resonant circuit is defined as:**
    A) $\omega_0 L / R$
    B) $1 / (\omega_0 C R)$
    C) $\omega_0 / \Delta \omega$ (Bandwidth)
    D) All of the above
    **Answer: D**

96. **Bandwidth of a resonant circuit is the range of frequencies for which the power is greater than or equal to:**
    A) Half the maximum power
    B) Maximum power
    C) Double the maximum power
    D) Zero
    **Answer: A**

97. **The relation between Bandwidth (BW), Resonant frequency ($f_0$), and Quality Factor (Q) is:**
    A) $BW = f_0 \times Q$
    B) $BW = f_0 / Q$
    C) $BW = Q / f_0$
    D) $BW = 2f_0 Q$
    **Answer: B**

98. **In a parallel resonant circuit (ideal), the impedance at resonance is:**
    A) Zero
    B) $R$
    C) Infinite
    D) $1/R$
    **Answer: C**

99. **A parallel resonant circuit is also called:**
    A) Acceptor circuit
    B) Rejector circuit
    C) Filter circuit
    D) Amplifier circuit
    **Answer: B**

100. **Self-inductance is defined as the property of a coil due to which it:**
    A) Opposes change in current
    B) Opposes change in voltage
    C) Enhances current
    D) Stores charge
    **Answer: A**

101. **The unit of Mutual Inductance is:**
    A) Henry
    B) Farad
    C) Ohm
    D) Ampere
    **Answer: A**

102. **Mutual inductance occurs when:**
    A) Two coils are far apart
    B) Current in one coil changes and induces voltage in another
    C) Current is constant
    D) Coils are short-circuited
    **Answer: B**

103. **The coefficient of coupling (k) ranges between:**
    A) 0 and 1
    B) 1 and 10
    C) -1 and 1
    D) 0 and $\infty$
    **Answer: A**

104. **If $k=1$, the coupling is called:**
    A) Loose coupling
    B) Tight coupling
    C) Critical coupling
    D) No coupling
    **Answer: B**

105. **Dot convention in coupled circuits determines:**
    A) Polarity of induced voltage
    B) Magnitude of induced voltage
    C) Power factor
    D) Resistance value
    **Answer: A**

106. **Problem: Two coils have self-inductances $L_1 = 2H$ and $L_2 = 8H$. The mutual inductance is $M=2H$. The coefficient of coupling is:**
    A) 1
    B) 0.5
    C) 0.25
    D) 0.75
    **Answer: B** ($k = M / \sqrt{L_1 L_2} = 2 / \sqrt{16} = 0.5$)

107. **In an ideal transformer, the relation between primary ($V_1$) and secondary ($V_2$) voltage and turns ($N_1, N_2$) is:**
    A) $V_1/V_2 = (N_1/N_2)^2$
    B) $V_1/V_2 = N_1/N_2$
    C) $V_1/V_2 = N_2/N_1$
    D) $V_1 V_2 = N_1 N_2$
    **Answer: B**

108. **In an ideal transformer, the power in the primary side is ______ the power in the secondary side.**
    A) Greater than
    B) Less than
    C) Equal to
    D) Double
    **Answer: C**

109. **Impedance transformation property of a transformer allows impedance on the secondary side to be referred to primary side by multiplying with:**
    A) $N_1/N_2$
    B) $(N_1/N_2)^2$
    C) $(N_2/N_1)^2$
    D) $N_2/N_1$
    **Answer: B**

110. **A linear transformer differs from an ideal transformer because it has:**
    A) Infinite permeability
    B) Finite winding resistance and leakage flux
    C) Unity coupling coefficient
    D) Zero core losses
    **Answer: B**

111. **Selectivity of a resonant circuit is better if:**
    A) Q is low
    B) Q is high
    C) Bandwidth is large
    D) R is high
    **Answer: B**

112. **For a series RLC circuit, if the frequency is below the resonant frequency, the circuit behaves as:**
    A) Resistive
    B) Inductive
    C) Capacitive
    D) Open circuit
    **Answer: C** ($X_C > X_L$)

113. **For a series RLC circuit, if the frequency is above the resonant frequency, the circuit behaves as:**
    A) Resistive
    B) Inductive
    C) Capacitive
    D) Short circuit
    **Answer: B** ($X_L > X_C$)

114. **The effective inductance of two coils in series aiding connection ($L_1, L_2, M$) is:**
    A) $L_1 + L_2$
    B) $L_1 + L_2 + 2M$
    C) $L_1 + L_2 - 2M$
    D) $L_1 L_2$
    **Answer: B**

115. **The effective inductance of two coils in series opposing connection is:**
    A) $L_1 + L_2$
    B) $L_1 + L_2 + 2M$
    C) $L_1 + L_2 - 2M$
    D) $L_1 L_2$
    **Answer: C**

---

## **Section 6: Three Phase Circuits**

116. **In a three-phase system, the phase sequence is important for:**
    A) Calculating power
    B) Determining the direction of rotation of motors
    C) Calculating current
    D) Finding resistance
    **Answer: B**

117. **The standard phase sequence is usually:**
    A) A-B-C
    B) B-A-C
    C) C-B-A
    D) A-C-B
    **Answer: A**

118. **In a balanced 3-phase Star connection, the line voltage ($V_L$) and Phase voltage ($V_{ph}$) are related by:**
    A) $V_L = V_{ph}$
    B) $V_L = \sqrt{3} V_{ph}$
    C) $V_L = 3 V_{ph}$
    D) $V_L = V_{ph} / \sqrt{3}$
    **Answer: B**

119. **In a balanced 3-phase Star connection, the line current ($I_L$) and Phase current ($I_{ph}$) are related by:**
    A) $I_L = I_{ph}$
    B) $I_L = \sqrt{3} I_{ph}$
    C) $I_L = 3 I_{ph}$
    D) $I_L = I_{ph} / \sqrt{3}$
    **Answer: A**

120. **In a balanced 3-phase Delta connection, the relation between Line Voltage and Phase Voltage is:**
    A) $V_L = \sqrt{3} V_{ph}$
    B) $V_L = V_{ph}$
    C) $V_L = V_{ph} / \sqrt{3}$
    D) $V_L = 3 V_{ph}$
    **Answer: B**

121. **In a balanced 3-phase Delta connection, the relation between Line Current and Phase Current is:**
    A) $I_L = I_{ph}$
    B) $I_L = \sqrt{3} I_{ph}$
    C) $I_L = I_{ph} / \sqrt{3}$
    D) $I_L = 3 I_{ph}$
    **Answer: B**

122. **The total power in a 3-phase balanced circuit is:**
    A) $3 V_{ph} I_{ph} \cos\phi$
    B) $\sqrt{3} V_L I_L \cos\phi$
    C) Both A and B
    D) None of the above
    **Answer: C**

123. **The phasor sum of EMFs in a balanced 3-phase system is:**
    A) $3 E$
    B) $\sqrt{3} E$
    C) Zero
    D) Infinite
    **Answer: C**

124. **A Star-Star connection is mainly used for:**
    A) High current applications
    B) High voltage applications
    C) Low voltage applications
    D) Motor loads
    **Answer: B**

125. **A Delta-Delta connection is mainly used for:**
    A) High voltage transmission
    B) High current loads / Distribution
    C) Instrument transformers
    D) Stepping up voltage
    **Answer: B**

126. **A Star-Delta transformer connection is used to:**
    A) Step up voltage
    B) Step down voltage
    C) Both A and B depending on the side
    D) Isolate the circuit
    **Answer: C**

127. **The neutral point in a Star connection is used to:**
    A) Provide a return path for unbalanced current
    B) Provide single-phase voltage
    C) Protect equipment (grounding)
    D) All of the above
    **Answer: D**

128. **In a 3-wire system, which connection is used?**
    A) Star with neutral
    B) Delta
    C) Open Delta
    D) Both B and C
    **Answer: D**

129. **Two-wattmeter method is used to measure power in:**
    A) Star connected load only
    B) Delta connected load only
    C) Balanced and unbalanced 3-phase 3-wire loads
    D) Only unbalanced loads
    **Answer: C**

130. **In the Two-wattmeter method, if the load is purely resistive, the readings of W1 and W2 are:**
    A) Equal
    B) W1 > W2
    C) W2 > W1
    D) One is zero
    **Answer: A**

131. **In Two-wattmeter method, if $W_1 = W_2$, the power factor is:**
    A) 0
    B) 0.5
    C) 1
    D) 0.866
    **Answer: C**

132. **In Two-wattmeter method, if one wattmeter reads zero, the power factor is:**
    A) 1
    B) 0.5
    C) 0
    D) 0.866
    **Answer: B**

133. **The reactive power in a 3-phase circuit can be measured by:**
    A) One wattmeter method
    B) Two wattmeter method
    C) Three wattmeter method
    D) All of the above
    **Answer: D**

134. **Problem: A balanced star load has impedance $10\Omega$ per phase and line voltage is 400V. The phase current is:**
    A) 40A
    B) 23.09A
    C) 14.14A
    D) 69.28A
    **Answer: B** ($V_{ph} = 400/\sqrt{3} = 230.94V$. $I_{ph} = 230.94/10 = 23.09A$)

135. **Problem: In a Delta load, if Line Current is 17.32A, the Phase Current is:**
    A) 30A
    B) 10A
    C) 15A
    D) 5.77A
    **Answer: B** ($I_{ph} = I_L / \sqrt{3} = 17.32 / 1.732 = 10A$)

---

## **Section 7: Mixed Problem Solving & Advanced Concepts**

136. **Problem: Calculate the equivalent resistance between terminals A and B for an infinite ladder network of 1Ω resistors.**
    A) 1 $\Omega$
    B) 2 $\Omega$
    C) $(1+\sqrt{5})/2$ $\Omega$
    D) $\infty$
    **Answer: C** (Golden ratio solution)

137. **Concept: A circuit has a capacitor and an inductor in series. At $t=0$, a DC source is connected. The initial current is:**
    A) Maximum
    B) Zero
    C) Infinite
    D) Depends on voltage
    **Answer: B** (Inductor opposes instantaneous current change)

138. **Problem: A voltage source $V_s = 10\sin(\omega t)$ is applied to a resistor of 5Ω. The power dissipated is:**
    A) 20W
    B) 10W
    C) 5W
    D) 2W
    **Answer: B** (RMS voltage is $10/\sqrt{2}$. Power = $V_{rms}^2 / R = 50 / 5 = 10W$)

139. **In a practical current source, the internal resistance is:**
    A) Very high
    B) Very low
    C) Zero
    D) Infinite
    **Answer: A**

140. **The condition for a circuit to be "Dual" of another is:**
    A) Elements are identical
    B) Series becomes Parallel, Voltage becomes Current
    C) Resistors become Capacitors
    D) Structure is reversed
    **Answer: B**

141. **Which theorem is used to analyze circuits with only one non-linear element?**
    A) Superposition
    B) Thevenin's
    C) Norton's
    D) Reciprocity
    **Answer: B** (Reduce the linear part to a source and resistance).

142. **Problem: A 12V battery has an internal resistance of 0.5Ω. What is the maximum power it can deliver to an external load?**
    A) 144W
    B) 72W
    C) 36W
    D) 24W
    **Answer: B** (Max power when $R_L = R_{int} = 0.5\Omega$. Total R = 1Ω. $I = 12A$. $P = I^2 R_L = 144 \times 0.5 = 72W$)

143. **Problem: In a Supermesh, which equation is missing that must be added?**
    A) KVL of the outer loop
    B) KCL at the node connecting the meshes
    C) Ohm's law
    D) KVL of individual meshes
    **Answer: B** (The constraint equation from the current source).

144. **If a circuit is purely capacitive, the Reactive Power (Q) is:**
    A) Positive
    B) Negative
    C) Zero
    D) Complex
    **Answer: B** (Capacitors deliver reactive power, Q is negative by convention).

145. **Power Factor correction involves adding a ______ in parallel with the load.**
    A) Resistor
    B) Inductor
    C) Capacitor
    D) Voltage source
    **Answer: C**

146. **Why is the 3-phase system preferred over single-phase for power transmission?**
    A) Less copper required for same power
    B) Constant instantaneous power
    C) Smoother vibration in machines
    D) All of the above
    **Answer: D**

147. **Problem: A transformer has 200 turns on primary and 50 turns on secondary. If 100V AC is applied, the secondary voltage is:**
    A) 400V
    B) 25V
    C) 100V
    D) 50V
    **Answer: B** ($V_2 = V_1 \times N_2/N_1 = 100 \times 50/200 = 25V$)

148. **In an RLC circuit, if the resistance $R \to 0$, the bandwidth:**
    A) $\to 0$
    B) $\to \infty$
    C) Remains constant
    D) Increases linearly
    **Answer: A** (Q becomes infinite, selectivity becomes infinite, bandwidth becomes zero).

149. **Tellegen's Theorem verifies the solution of a circuit by checking:**
    A) Voltage balance
    B) Current balance
    C) Power balance (Sum of powers = 0)
    D) Energy balance
    **Answer: C**

150. **The unit of Complex Power (S) is:**
    A) Watt
    B) VAR
    C) VA (Volt-Ampere)
    D) Joule
    **Answer: C**

151. **Concept: In a node, if current entering is $I_1$ and leaving is $I_2$, the nodal equation is:**
    A) $I_1 + I_2 = 0$
    B) $I_1 - I_2 = 0$
    C) $I_1 \times I_2 = 0$
    D) $V = I_1 R$
    **Answer: B** (Sum of currents entering = 0)

152. **A series RLC circuit with $R=10, L=0.1H, C=10\mu F$ is connected to 100V DC. The final voltage across the capacitor is:**
    A) 0V
    B) 100V
    C) 10V
    D) Infinite
    **Answer: B** (Capacitor acts as open circuit to DC at steady state).

153. **For the same circuit above, the final current through the inductor is:**
    A) 10A
    B) 0A
    C) 100A
    D) 1A
    **Answer: A** (Inductor acts as short circuit to DC at steady state. $I = V/R = 100/10 = 10A$).

154. **Problem: Three identical impedances are connected in Star to a 3-phase supply. If they are reconnected in Delta to the same supply, the power consumed will be:**
    A) Same
    B) Three times
    C) One-third
    D) $\sqrt{3}$ times
    **Answer: B** ($P_{Delta} = 3 \times P_{Star}$ for same impedance and line voltage).

155. **The RMS value of a triangular wave with peak $V_m$ is:**
    A) $V_m$
    B) $V_m / \sqrt{2}$
    C) $V_m / \sqrt{3}$
    D) $V_m / 2$
    **Answer: C**

156. **Which component is responsible for "Hysteresis Loss" in AC circuits?**
    A) Resistor
    B) Capacitor
    C) Inductor (Core)
    D) Diode
    **Answer: C**

157. **Problem: Find the Thevenin resistance if $V_{th} = 20V$ and $I_{sc} = 5A$.**
    A) 100 $\Omega$
    B) 25 $\Omega$
    C) 4 $\Omega$
    D) 0.25 $\Omega$
    **Answer: C** ($R_{th} = V_{th}/I_{sc} = 20/5 = 4\Omega$)

158. **A voltage $v(t) = 100\sin(314t + 30^\circ)$ has a frequency of:**
    A) 314 Hz
    B) 50 Hz
    C) 100 Hz
    D) 60 Hz
    **Answer: B** ($\omega = 314 \approx 100\pi$, $f = \omega/2\pi = 50Hz$)

159. **In a network, if the loop currents are defined, the mutual resistance between loop $i$ and loop $j$ is positive if:**
    A) Loop currents are opposite
    B) Loop currents are same direction in the common branch
    C) There is no common branch
    D) It is always negative
    **Answer: B**

160. **A "Sink" in a circuit diagram usually represents:**
    A) A voltage source
    B) A node
    C) Ground / Reference
    D) A switch
    **Answer: C**

161. **The Fourier Series is used to analyze:**
    A) DC circuits
    B) Sinusoidal AC circuits
    C) Non-sinusoidal periodic waveforms
    D) Transient waveforms
    **Answer: C**

162. **Problem: A load of 11kW has a power factor of 0.8 lagging. The Reactive Power (Q) is:**
    A) 11 kVAR
    B) 8.25 kVAR
    C) 6.6 kVAR
    D) 13.75 kVAR
    **Answer: B** ($\cos\phi = 0.8 \to \sin\phi = 0.6$. $S = P/\cos\phi = 13.75kVA$. $Q = S \sin\phi = 13.75 \times 0.6 = 8.25kVAR$)

163. **What happens if the primary of a transformer is connected to a DC source?**
    A) Transformer works normally
    B) Primary winding burns out
    C) Secondary voltage is high
    D) No effect
    **Answer: B** (Inductive reactance $X_L = 2\pi f L$. For DC $f=0$, so $X_L=0$. Only winding resistance remains, causing huge current).

164. **The "Polar form" of a complex number $A + jB$ is:**
    A) $\sqrt{A^2+B^2} \angle \tan^{-1}(B/A)$
    B) $\sqrt{A^2-B^2} \angle \tan^{-1}(A/B)$
    C) $A \angle B$
    D) $\sqrt{AB} \angle 90$
    **Answer: A**

165. **In an AC circuit, if the current leads the voltage, the load is:**
    A) Inductive
    B) Resistive
    C) Capacitive
    D) Combination of R and L
    **Answer: C**

166. **Problem: Three capacitors of 3$\mu$F, 6$\mu$F, and 9$\mu$F are connected in series. The total capacitance is:**
    A) 18 $\mu$F
    B) 1.64 $\mu$F
    C) 6 $\mu$F
    D) 5.14 $\mu$F
    **Answer: B** ($1/C_{eq} = 1/3 + 1/6 + 1/9 = 11/18$. $C_{eq} = 18/11 \approx 1.64\mu F$)

167. **The quality factor $Q$ is a measure of:**
    A) Power loss
    B) Energy stored per cycle / Energy dissipated per cycle
    C) Bandwidth only
    D) Resistance value
    **Answer: B**

168. **In a coupled circuit, the "Leakage Flux" is responsible for:**
    A) Perfect coupling
    B) Loss in coupling efficiency
    C) Increased inductance
    D) Zero mutual inductance
    **Answer: B**

169. **Problem: A 3-phase, 400V supply feeds a balanced delta load of 30$\Omega$ per phase. The line current is:**
    A) 13.33 A
    B) 23.09 A
    C) 40 A
    D) 69.28 A
    **Answer: B** ($V_{ph} = V_L = 400$. $I_{ph} = 400/30 = 13.33A$. $I_L = \sqrt{3} I_{ph} = 1.732 \times 13.33 \approx 23.09A$)

170. **The phase angle difference between line voltage and phase voltage in a Star connection is:**
    A) $0^\circ$
    B) $30^\circ$
    C) $90^\circ$
    D) $120^\circ$
    **Answer: B** (Line voltage leads respective phase voltage by 30 degrees).

171. **In a series RLC circuit, the voltage across L and C are equal at resonance. This is because:**
     A) $X_L = X_C$
     B) Current is maximum
     C) They are 180 degrees out of phase
     D) All of the above
     **Answer: D**

172. **A circuit has a voltage source $V=100V$ and two resistors $R_1=10\Omega, R_2=10\Omega$ in series. If $R_2$ is open-circuited, the voltage across the open terminal is:**
     A) 0V
     B) 50V
     C) 100V
     D) 200V
     **Answer: C** (No current means no drop across $R_1$).

173. **The purpose of using a capacitor in a single-phase induction motor is to:**
     A) Improve power factor
     B) Create a phase shift for starting torque
     C) Reduce noise
     D) Store energy
     **Answer: B**

174. **If a resistor is color coded: Brown, Black, Red, Gold. What is its value?**
     A) 100 $\Omega$ $\pm$ 5%
     B) 1 k$\Omega$ $\pm$ 5%
     C) 10 $\Omega$ $\pm$ 5%
     D) 1000 $\Omega$ $\pm$ 10%
     **Answer: B** ($1, 0, \times 10^2, \pm 5\%$)

175. **The "Inverse Laplace Transform" method is used to convert:**
     A) Time domain to Frequency domain
     B) Frequency domain to Time domain
     C) Phasor to Vector
     D) AC to DC
     **Answer: B**

176. **A "Reflected Impedance" in a transformer refers to:**
     A) Impedance of the primary coil
     B) Impedance of the secondary coil as seen from the primary
     C) Impedance of the source
     D) Loss resistance
     **Answer: B**

177. **Problem: A load draws 2kW at 0.8 pf lagging. The capacitance required to correct the power factor to unity (Assume 230V, 50Hz) is roughly:**
     A) 50 $\mu$F
     B) 30 $\mu$F
     C) 10 $\mu$F
     D) 100 $\mu$F
     *(Calculation: $P=2000$. $\phi_1 = \cos^{-1}0.8 = 36.8^\circ$. $Q_1 = P \tan \phi = 1500 VAR$. Unity pf means $Q_{new}=0$. Capacitor must supply 1500 VAR. $Q_C = V^2 \omega C$. $C = 1500 / (230^2 \times 314) \approx 90 \mu F$. Roughly 100uF)*
     **Answer: D**

178. **The phenomenon of "Skin Effect" in AC conductors causes:**
     A) Increase in effective resistance
     B) Increase in inductance
     C) Decrease in capacitance
     D) Uniform current distribution
     **Answer: A**

179. **A network function $H(s)$ is stable if all poles lie in the:**
     A) Right half of s-plane
     B) Left half of s-plane
     C) On the imaginary axis
     D) Origin
     **Answer: B**

180. **In a 3-phase 4-wire system, the current in the neutral wire is zero if:**
     A) The load is unbalanced
     B) The load is balanced
     C) One phase is open
     D) It is never zero
     **Answer: B**

181. **Problem: A voltage source $50\angle0^\circ$ V is in series with $Z = 3+j4 \Omega$. The maximum power transferred to a load $Z_L$ will be:**
     A) 150W
     B) 250W
     C) 300W
     D) 125W
     **Answer: A** (For max power, $Z_L = 3-j4$. $R_{total} = 6\Omega$. $I = 50/6 = 8.33A$. $P = I^2 R_L = (8.33)^2 \times 3 \approx 208W$.)

182. **The degree of freedom for a circuit with 5 nodes and 8 branches is:**
     A) 4
     B) 13
     C) 3
     D) 8
     **Answer: A** ($b - n + 1 = 8 - 5 + 1 = 4$)

183. **A capacitor $C$ is charged to voltage $V_0$. At $t=0$, it is connected to an inductor $L$. The maximum current in the LC oscillation is:**
     A) $V_0 \sqrt{C/L}$
     B) $V_0 \sqrt{L/C}$
     C) $V_0 / L$
     D) $V_0 C$
     **Answer: A** (Energy balance: $0.5 C V_0^2 = 0.5 L I_{max}^2$)

184. **In a practical voltage source, the terminal voltage drops as load current increases because:**
     A) Internal resistance drops voltage
     B) Load resistance decreases
     C) Source EMF decreases
     D) Power factor reduces
     **Answer: A**

185. **Problem: Two wattmeters connected to measure power read -200W and +800W. The total power is:**
     A) 600W
     B) 1000W
     C) 200W
     D) 800W
     **Answer: A** (Algebraic sum: $800 + (-200) = 600W$)

186. **The inverse Laplace of $\frac{10}{s(s+2)}$ is:**
     A) $5(1-e^{-2t})$
     B) $10(1-e^{-2t})$
     C) $5 e^{-2t}$
     D) $10 e^{-2t}$
     **Answer: A** (Partial fractions: $\frac{10}{s(s+2)} = \frac{5}{s} - \frac{5}{s+2}$. Result $5 - 5e^{-2t}$).

187. **A transformer has 1:2 turns ratio. If 100V is applied, the current in secondary is 5A. The primary current is:**
     A) 2.5A
     B) 5A
     C) 10A
     D) 20A
     **Answer: C** ($I_1 V_1 = I_2 V_2 \to I_1 \times 100 = 5 \times 200 \to I_1 = 10A$).

188. **Which analysis method is best for a circuit containing only current sources?**
     A) Mesh
     B) Node
     C) Both are difficult
     D) Superposition
     **Answer: B**

189. **For a series R-L-C circuit, the half-power frequencies are also called:**
     A) Resonant frequencies
     B) Cut-off frequencies
     C) Corner frequencies
     D) Both B and C
     **Answer: D**

190. **Problem: A 20mH inductor is connected in series with a resistor of $100\Omega$. The time required for the current to reach 63.2% of its final value is:**
     A) 0.2 ms
     B) 0.2 s
     C) 5 ms
     D) 5 s
     **Answer: A** ($\tau = L/R = 20\times10^{-3}/100 = 0.2ms$).

191. **In an AC circuit, $V = 100\angle30^\circ$ and $I = 10\angle-30^\circ$. The nature of the circuit is:**
     A) Resistive
     B) Inductive
     C) Capacitive
     D) Cannot determine
     **Answer: B** (Voltage leads current. Hence Inductive).

192. **A circuit that separates different frequencies is called a:**
     A) Amplifier
     B) Filter
     C) Oscillator
     D) Rectifier
     **Answer: B**

193. **In a Delta connection, if one of the phases is open, it becomes an "Open Delta" or "V-connection". The capacity of the Open Delta bank is:**
     A) 100% of original
     B) 86.6% of original
     C) 57.7% of original
     D) 50% of original
     **Answer: C** ($\frac{1}{\sqrt{3}} \approx 0.577$).

194. **The "Node" in an electrical circuit is a point where:**
     A) Two or more elements meet
     B) Voltage is zero
     C) Current is zero
     D) Power is dissipated
     **Answer: A**

195. **In source transformation, a voltage source $V$ in series with resistance $R$ is equivalent to a current source:**
     A) $V/R$ in parallel with $R$
     B) $V \times R$ in parallel with $R$
     C) $V/R$ in series with $R$
     D) $V \times R$ in series with $R$
     **Answer: A**

196. **Problem: Find the current $I$ in a loop containing a 12V source, $2\Omega$ and $4\Omega$ resistors in series, assuming clockwise current direction.**
     A) 2A
     B) 3A
     C) 6A
     D) 12A
     **Answer: A** ($I = V/R_{total} = 12/6 = 2A$).

197. **The phasor diagram for a purely resistive circuit shows:**
     A) V leading I
     B) I leading V
     C) V and I in phase
     D) V and I 180 deg apart
     **Answer: C**

198. **In a complex circuit, a "Supernode" is formed between two nodes if they are connected by a:**
     A) Resistor
     B) Inductor
     C) Dependent Voltage Source
     D) Capacitor
     **Answer: C** (Also an independent voltage source).

199. **A circuit has a dependent voltage source. To find Thevenin resistance, which method is invalid?**
    A) Open circuit voltage / Short circuit current
    B) Deactivating independent sources and applying a test source
    C) Deactivating all sources (independent and dependent) and calculating equivalent resistance
    D) Using experimental measurement
    **Answer: C** (Dependent sources cannot be deactivated).

200. **In a Star-Connected 3-phase system, the line voltage $V_{RY}$ leads the phase voltage $V_{RN}$ by:**
    A) $-30^\circ$
    B) $30^\circ$
    C) $90^\circ$
    D) $0^\circ$
    **Answer: B**

---

## **Section 8: Critical Thinking & Diagram Based Questions (Nodal & Load Flow)**

**201. Diagram Description (Nodal):**
*A circuit has a Reference Node (Ground). Node 1 is connected to Node 2 via a $2\Omega$ resistor. Node 1 is also connected to Ground via a $4\Omega$ resistor. Node 2 is connected to Ground via a $2\Omega$ resistor. A current source of $10A$ is feeding into Node 1.*

**Question:** What is the conductance matrix $[G]$ for this circuit?
A) $\begin{bmatrix} 0.75 & -0.5 \\ -0.5 & 1.0 \end{bmatrix}$
B) $\begin{bmatrix} 0.5 & -0.5 \\ -0.5 & 1.0 \end{bmatrix}$
C) $\begin{bmatrix} 1.0 & -0.5 \\ -0.5 & 0.75 \end{bmatrix}$
D) $\begin{bmatrix} 0.75 & -0.5 \\ -0.5 & 0.75 \end{bmatrix}$
**Answer: A**
*Explanation:* $G_{11} = 1/4 + 1/2 = 0.75S$. $G_{22} = 1/2 + 1/2 = 1.0S$. Off-diagonal is $-1/2 = -0.5S$.

**202. Diagram Description (Nodal):**
*A circuit contains Node A and Node B. A Voltage Source of $20V$ is connected between Node A and the Reference Node (Positive at A). A Resistor of $5\Omega$ is connected between Node A and Node B. A Resistor of $10\Omega$ is connected between Node B and the Reference Node.*

**Question:** If Nodal Analysis is performed, what is the voltage at Node B?
A) 20V
B) 13.33V
C) 10V
D) 6.67V
**Answer: B**
*Explanation:* $V_A = 20V$. Voltage divides across $5\Omega$ and $10\Omega$. $V_B = 20 \times \frac{10}{10+5} = 13.33V$.

**203. Diagram Description (Supernode):**
*A circuit has Node 1 and Node 2. An ideal Voltage Source of $12V$ is connected between Node 1 and Node 2 (Positive at Node 2). There is no impedance in series with the voltage source. Node 1 is connected to Ground via a $2\Omega$ resistor. Node 2 is connected to Ground via a $4\Omega$ resistor. A current source of $5A$ enters Node 1.*

**Question:** Which of the following equations correctly represents the constraint equation for the Supernode?
A) $V_1 - V_2 = 12$
B) $V_2 - V_1 = 12$
C) $V_1 + V_2 = 12$
D) $V_2 = 12$
**Answer: B**
*Explanation:* Since positive terminal is at Node 2, $V_2$ is higher than $V_1$ by 12V.

**204. Conceptual Question (Dependent Source):**
*A circuit contains a Voltage Controlled Current Source (VCCS) connected at Node X. The VCCS provides a current $I = 0.5 V_{Y}$, where $V_Y$ is the voltage at Node Y. The current direction is leaving Node X.*

**Question:** How does this affect the Nodal Conductance Matrix?
A) It adds a constant current to the Current Vector.
B) It modifies the off-diagonal elements and diagonal elements of the Conductance Matrix.
C) It has no effect on the matrix; it is treated as an independent source.
D) It makes Nodal Analysis invalid.
**Answer: B**
*Explanation:* Dependent sources couple the equations of two nodes.

**205. Load Flow Analysis:**
*In Load Flow Studies, the system is assumed to be operating under which condition?*
A) Transient state
B) Sub-transient state
C) Steady-state balanced condition
D) Fault condition
**Answer: C**

**206. Load Flow Analysis:**
*A "PQ Bus" (Load Bus) is one where:*
A) Real Power (P) and Reactive Power (Q) are unknown.
B) Voltage Magnitude (V) and Voltage Angle ($\delta$) are known.
C) Real Power (P) and Reactive Power (Q) are specified.
D) Real Power (P) and Voltage Magnitude (V) are specified.
**Answer: C**

**207. Load Flow Analysis:**
*Which type of bus is used as the reference node for the entire power system network?*
A) PV Bus (Generator Bus)
B) PQ Bus (Load Bus)
C) Slack Bus (Swing Bus)
D) Z Bus
**Answer: C**

**208. Diagram Description (Y-Bus):**
*A simple 3-bus power system is shown. Bus 1 is connected to Bus 2 via a line with admittance $y_{12} = -j10 S$. Bus 2 is connected to Bus 3 via a line with admittance $y_{23} = -j5 S$. Bus 1 is connected to Bus 3 via a line with admittance $y_{13} = -j10 S$. There are no shunt admittances.*

**Question:** What is the diagonal element $Y_{22}$ of the Bus Admittance Matrix ($Y_{BUS}$)?
A) $-j15 S$
B) $j25 S$
C) $-j25 S$
D) $j15 S$
**Answer: A**
*Explanation:* Sum of admittances connected to Bus 2: $y_{21} + y_{23} = -j10 + -j5 = -j15$.

**209. Problem Solving (Slack Bus):**
*In a load flow solution, the total calculated system losses are 10 MW. The sum of all load demands is 100 MW, and the sum of all generator outputs (excluding the slack bus) is 80 MW.*

**Question:** What is the power generation at the Slack Bus?
A) 10 MW
B) 20 MW
C) 30 MW
D) 90 MW
**Answer: C**
*Explanation:* Total Gen = Total Load + Losses. $P_{slack} + 80 = 100 + 10$. $P_{slack} = 30 MW$.

**210. Y-Bus Construction:**
*Which of the following properties is TRUE for the Bus Admittance Matrix ($Y_{BUS}$)?*
A) It is a non-singular matrix.
B) It is a sparse matrix.
C) It is always a diagonal matrix.
D) It is independent of the network topology.
**Answer: B**
*Explanation:* In large power systems, each bus is connected to only a few other buses, creating many zero elements (sparsity).

**211. Conceptual Diagram:**
*A generator is connected to an infinite bus through a transmission line. The infinite bus voltage is $V$. The generator internal voltage is $E$. The line reactance is $X$.*

**Question:** If the generator increases its excitation (increases $|E|$) while keeping the mechanical power input constant, what happens to the reactive power output $Q$?
A) $Q$ increases (Generator supplies vars).
B) $Q$ decreases (Generator absorbs vars).
C) $Q$ remains constant.
D) Generator trips.
**Answer: A**
*Explanation:* Increasing internal voltage magnitude increases reactive power export.

**212. Nodal vs Mesh in Power Systems:**
*Why is Nodal Analysis (Y-Bus) preferred over Mesh Analysis (Z-Bus) for Load Flow studies?*
A) Mesh analysis requires defining loops which are difficult to identify in large interconnected networks.
B) Nodal analysis works better with current sources.
C) The number of nodes is usually smaller than the number of loops in power grids.
D) Both A and C.
**Answer: D**

**213. Problem: Voltage Regulation**
*A load flow study shows that a remote load bus has a voltage magnitude of $0.90$ pu. The nominal voltage is $1.0$ pu.*

**Question:** To improve the voltage profile to $1.0$ pu, which device should be installed at the load bus?
A) Shunt Capacitor Bank
B) Shunt Reactor
C) Series Capacitor
D) Tap-changing transformer (decreasing tap)
**Answer: A**
*Explanation:* Low voltage indicates a shortage of reactive power. Shunt capacitors inject reactive power.

**214. Gauss-Seidel vs Newton-Raphson:**
*Which method is preferred for Load Flow analysis when the system size is very large and convergence speed is critical?*
A) Gauss-Seidel Method
B) Newton-Raphson Method
C) Backward-Forward Sweep
D) Trial and Error
**Answer: B**
*Explanation:* Newton-Raphson has faster convergence characteristics for large systems compared to Gauss-Seidel.

**215. Power Flow Equation:**
*The real power flow equation between two buses is approximately $P = \frac{|V_1||V_2|}{X} \sin(\delta_{12})$. This indicates that real power flow is primarily controlled by:*
A) Voltage Magnitude difference
B) Voltage Angle difference
C) Line Resistance
D) Frequency
**Answer: B**
"""
    
    doc = Document()
    
    # Set default font
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = 11

    for line in text_content.split('\n'):
        line = line.strip()
        if not line:
            continue
        
        # Main Title
        if line.startswith('# ') and not line.startswith('## '):
            doc.add_heading(line.replace('# ', ''), level=1)
        # Section Subtitles
        elif line.startswith('## '):
            doc.add_heading(line.replace('## ', ''), level=2)
        # Bold Answers or Questions
        elif line.startswith('**Answer:') or line.startswith('**Explanation:'):
            p = doc.add_paragraph()
            runner = p.add_run(line.replace('**', ''))
            runner.bold = True
        # Questions (usually start with a number)
        elif re.match(r'^\d+\.\s\*\*', line):
            p = doc.add_paragraph()
            clean_line = line.replace('**', '')
            runner = p.add_run(clean_line)
            runner.bold = True
        # Normal options or text
        else:
            doc.add_paragraph(line.replace('**', ''))

    # Save the document
    filename = "Electric_Circuit_Analysis_MCQ_Bank.docx"
    doc.save(filename)
    print(f"Success! Document saved as {filename}")

if __name__ == "__main__":
    create_mcq_bank()