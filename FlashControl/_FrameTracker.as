package  {
	
	import flash.display.MovieClip;
	import flash.events.Event;
	import flash.external.ExternalInterface;
	
	
	public class _FrameTracker extends MovieClip {
		private static var initialized: Boolean = false;
		
		public function _FrameTracker()
		{
			if (!initialized) {
				initialized = true;
				stage.root.addEventListener(Event.ENTER_FRAME, onEnterFrame, false);
			}
		}
		private static function onEnterFrame(e: Event)
		{
			ExternalInterface.call('onEnterFrame');
		}
	}
	
}
