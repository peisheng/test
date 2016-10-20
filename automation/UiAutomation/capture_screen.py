#!python3
# -*- coding: utf-8 -*-
import os
import time
import automation

def CaptureControl(c, path, up = False):
    if c.CaptureToImage(path):
        os.popen(path)
        automation.Logger.WriteLine('capture image: ' + path)
    else:
        automation.Logger.WriteLine('capture failed', automation.ConsoleColor.Yellow)
    if up:
        r = automation.GetRootControl()
        depth = 0
        name, ext = os.path.splitext(path)
        while True:
            c = c.GetParentControl()
            if not c or automation.ControlsAreSame(c, r):
                break
            depth += 1
            savePath = name + '_p' * depth + ext
            if c.CaptureToImage(savePath):
                automation.Logger.WriteLine('capture image: ' + savePath)

def main(args):
    if args.time > 0:
        time.sleep(args.time)
    start = time.clock()
    if args.active:
        c = automation.GetForegroundControl()
    elif args.cursor or args.up:
        c = automation.ControlFromCursor()
    elif args.fullscreen:
        c = automation.GetRootControl()
    CaptureControl(c, args.path, args.up)
    automation.Logger.WriteLine('cost time: {:.3f} s'.format(time.clock() - start))


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('-f', '--fullscreen', action='store_true', default = True,  help = 'capture full screen')
    parser.add_argument('-a', '--active', action='store_true', help = 'capture active window')
    parser.add_argument('-c', '--cursor', action='store_true', help = 'capture control under cursor')
    parser.add_argument('-u', '--up', action='store_true', help = 'capture control under cursor and up to its top window')
    parser.add_argument('-p', '--path', type = str, default = 'capture.png', help = 'save path')
    parser.add_argument('-t', '--time', type = int, default = 0, help = 'delay time')
    args = parser.parse_args()
    #automation.Logger.WriteLine(str(args))
    main(args)


