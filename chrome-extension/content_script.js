document.body.addEventListener('mouseover', evt => {
	let target = evt.target;
	while (target
		&& target.tagName.toLowerCase() !== 'a'
	) {
		target = target.parentElement;
	}
	if (!target) return;

	const url = target.href;
	if (!url.startsWith('file://')) return;

	const ext = (url.match(/[.]([^.]+)$/) || [])[1];
	if (!ext) return;

	const type = ({
		doc: 'word',
		docx: 'word',
		xls: 'excel',
		xlsx: 'excel',
		ppt: 'powerpoint',
		pptx: 'powerpoint',
	})[ext];
	if (!type) return;

	target.href = `ms-${type}:ofv|u|${url}`;
}, {
	capture: true,
});
