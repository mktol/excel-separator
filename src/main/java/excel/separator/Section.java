package excel.separator;

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

	public void setStart(final int start) {
		this.start = start;
	}

	public int getFinish() {
		return finish;
	}

	public void setFinish(final int finish) {
		this.finish = finish;
	}

	@Override
	public String toString() {
		return "Section{" +
				"start=" + start +
				", finish=" + finish +
				'}';
	}
}
