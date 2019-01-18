package excel.separator.entity;

public class Section {

	private int start;

	private int finish;

	public Section(final int start, final int finish) {
		this.start = start;
		this.finish = finish;
	}

	public int getStart() {
		return start;
	}

	public int getFinish() {
		return finish;
	}

	@Override
	public String toString() {
		return "Section{" +
				"start=" + start +
				", finish=" + finish +
				'}';
	}
}
