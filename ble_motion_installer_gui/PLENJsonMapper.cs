using System.Collections;
using System.Collections.Generic;

namespace PLEN.JSON {
	public class Main {
		public short slot;
		public string name;
		public List<Code> codes = new List<Code> ();
		public List<Frame> frames = new List<Frame>();
	}

	public class Code {
	}

	public class Frame {
		public int transition_time_ms;
		public List<Output> outputs = new List<Output>();
	}
	public class Output {
		public string device;
		public short value;
	}
}

namespace PLEN {
	public enum JointName {
		left_shoulder_pitch,
		left_thigh_yaw,
		left_shoulder_roll,
		left_elbow_roll,
		left_thigh_roll,
		left_thigh_pitch,
		left_knee_pitch,
		left_foot_pitch,
		left_foot_roll,
		right_shoulder_pitch,
		right_thigh_yaw,
		right_shoulder_roll,
		right_elbow_roll,
		right_thigh_roll,
		right_thigh_pitch,
		right_knee_pitch,
		right_foot_pitch,
		right_foot_roll
	}
}
