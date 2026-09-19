const workflowVideo = document.querySelector('[data-workflow-video] video');
const videoChapters = [...document.querySelectorAll('[data-video-time]')];
const videoFrame = document.querySelector('.videoFrame');
const videoPlay = document.querySelector('[data-video-play]');

function playWorkflowVideo() {
  if (!workflowVideo) return;
  workflowVideo.controls = true;
  workflowVideo.play();
}

function seekVideo(seconds) {
  if (!workflowVideo || !Number.isFinite(workflowVideo.duration)) return;
  workflowVideo.currentTime = Math.min(workflowVideo.duration, Math.max(0, workflowVideo.currentTime + seconds));
  workflowVideo.dispatchEvent(new Event('timeupdate'));
}

videoChapters.forEach((chapter) => {
  chapter.addEventListener('click', () => {
    if (!workflowVideo) return;
    workflowVideo.currentTime = Number(chapter.dataset.videoTime) || 0;
    workflowVideo.dispatchEvent(new Event('timeupdate'));
    playWorkflowVideo();
  });
});

videoPlay?.addEventListener('click', () => {
  playWorkflowVideo();
});

workflowVideo?.addEventListener('play', () => {
  videoFrame?.classList.add('hasStarted');
});

workflowVideo?.addEventListener('ended', () => {
  videoFrame?.classList.remove('hasStarted');
});

let videoFrameHovered = false;
videoFrame?.addEventListener('mouseenter', () => { videoFrameHovered = true; });
videoFrame?.addEventListener('mouseleave', () => { videoFrameHovered = false; });

document.addEventListener('keydown', (event) => {
  if (!workflowVideo || (!videoFrameHovered && !videoFrame?.contains(document.activeElement))) return;
  if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
  event.preventDefault();
  seekVideo(event.key === 'ArrowLeft' ? -10 : 10);
}, true);

workflowVideo?.addEventListener('timeupdate', () => {
  let activeChapter = videoChapters[0];
  videoChapters.forEach((chapter) => {
    if (workflowVideo.currentTime >= Number(chapter.dataset.videoTime)) activeChapter = chapter;
  });
  videoChapters.forEach((chapter) => chapter.classList.toggle('isActive', chapter === activeChapter));
});
