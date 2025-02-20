import sys
import os
import cv2

def extract_timestamps(timeline_file):
    with open(timeline_file, 'r') as f:
        lines = f.readlines()
    
    timestamps = []
    for i in range(1, len(lines), 2):
        start, end = lines[i].strip().split(' --> ')
        timestamps.append((start, end))
    
    return timestamps

def capture_frames(video_file, timestamps, output_dir_qian, output_dir_hou, video_index, total_videos):
    video = cv2.VideoCapture(video_file)
    fps = video.get(cv2.CAP_PROP_FPS)  # 获取视频的帧率
    total_timestamps = len(timestamps)
    
    print(f"总共导入了{total_videos}个视频 正在提取第{video_index}视频 {os.path.basename(video_file)} 总共要提取{total_timestamps}轴")
    print("-" * 50)
    
    qian_count = 0
    hou_count = 0
    
    # 根据总时间轴数量确定补0的位数
    zero_padding = len(str(total_timestamps))

    # 控制截图的参数
    # 选择模式和数值
    adjustment_mode = {
        'start': 'advance_seconds',  # 可选值: 'advance_seconds', 'delay_seconds', 'advance_frames', 'delay_frames'
        'end': 'advance_frames'       # 可选值: 'advance_seconds', 'delay_seconds', 'advance_frames', 'delay_frames'
    }
    adjustment_values = {
        'start_seconds': 1,  # `start` 时间轴的秒数调整
        'start_frames': 10,  # `start` 时间轴的帧数调整
        'end_seconds': 2,    # `end` 时间轴的秒数调整
        'end_frames': 10     # `end` 时间轴的帧数调整
    }
    
    for i, (start, end) in enumerate(timestamps, start=1):
        # 处理前时间轴的调整
        start_time_ms = convert_to_ms(start)
        if adjustment_mode['start'] == 'advance_seconds':
            start_time_ms -= adjustment_values['start_seconds'] * 1000  # 提前秒数
        elif adjustment_mode['start'] == 'delay_seconds':
            start_time_ms += adjustment_values['start_seconds'] * 1000  # 延后秒数
        elif adjustment_mode['start'] == 'advance_frames':
            start_time_ms -= adjustment_values['start_frames'] * 1000 / fps  # 提前帧数
        elif adjustment_mode['start'] == 'delay_frames':
            start_time_ms += adjustment_values['start_frames'] * 1000 / fps  # 延后帧数

        adjusted_start = convert_to_timestamp(start_time_ms)

        start_frame = int(start_time_ms * fps / 1000)

        # 处理后时间轴的调整
        end_time_ms = convert_to_ms(end)
        if adjustment_mode['end'] == 'advance_seconds':
            end_time_ms -= adjustment_values['end_seconds'] * 1000  # 提前秒数
        elif adjustment_mode['end'] == 'delay_seconds':
            end_time_ms += adjustment_values['end_seconds'] * 1000  # 延后秒数
        elif adjustment_mode['end'] == 'advance_frames':
            end_time_ms -= adjustment_values['end_frames'] * 1000 / fps  # 提前帧数
        elif adjustment_mode['end'] == 'delay_frames':
            end_time_ms += adjustment_values['end_frames'] * 1000 / fps  # 延后帧数

        adjusted_end = convert_to_timestamp(end_time_ms)

        end_frame = int(end_time_ms * fps / 1000)
        
        video.set(cv2.CAP_PROP_POS_FRAMES, start_frame)
        ret, frame_start = video.read()
        
        video.set(cv2.CAP_PROP_POS_FRAMES, end_frame)
        ret, frame_end = video.read()
        
        if ret:
            # 处理并保存截图
            cv2.rectangle(frame_start, (0, 0), (frame_start.shape[1], 40), (0, 0, 0), -1)
            text_start = f"Timestamp: {adjusted_start} (Aqian)  |  {i}/{total_timestamps}"
            cv2.putText(frame_start, text_start, (10, 25), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 2)
            
            cv2.rectangle(frame_end, (0, 0), (frame_end.shape[1], 40), (0, 0, 0), -1)
            text_end = f"Timestamp: {adjusted_end} (Bhou)  |  {i}/{total_timestamps}"
            cv2.putText(frame_end, text_end, (10, 25), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 2)
            
            output_path_qian = os.path.join(output_dir_qian, f"{i:0{zero_padding}d}_Aqian_{adjusted_start.replace(':', '-')}.jpg")
            cv2.imwrite(output_path_qian, frame_start)
            qian_count += 1
            
            output_path_hou = os.path.join(output_dir_hou, f"{i:0{zero_padding}d}_Bhou_{adjusted_end.replace(':', '-')}.jpg")
            cv2.imwrite(output_path_hou, frame_end)
            hou_count += 1
            
            print(f"总共导入了{total_videos}个视频 正在提取第{video_index}视频 {os.path.basename(video_file)} 提取{i}轴，前时间轴: {start}，后时间轴: {end}，目前提取进度{i}/{total_timestamps}")
        
        print("-" * 30)
    
    video.release()
    
    print(f"第{video_index}视频处理完成!")
    print(f"总共提取了 {qian_count + hou_count} 张截图")
    print(f"Aqian 文件夹: {qian_count} 张")
    print(f"Bhou 文件夹: {hou_count} 张")
    print("-" * 30)
    
    return qian_count, hou_count

def convert_to_ms(timestamp):
    h, m, s = timestamp.split(':')
    s, ms = s.split(',')
    return int(h) * 3600000 + int(m) * 60000 + int(s) * 1000 + int(ms)

def convert_to_timestamp(ms):
    h = ms // 3600000
    ms %= 3600000
    m = ms // 60000
    ms %= 60000
    s = ms // 1000
    ms %= 1000
    return f"{h:02}:{m:02}:{s:02},{ms // 10:02}"

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("请将字幕文件拖拽到此脚本上")
        input("按回车键退出...")
        sys.exit(1)
    
    total_videos = len(sys.argv) - 1
    total_qian_count = 0
    total_hou_count = 0
    video_results = []
    
    for video_index, timeline_file in enumerate(sys.argv[1:], start=1):
        video_file = os.path.splitext(timeline_file)[0] + '.mp4'
        
        if not os.path.exists(video_file):
            print(f"找不到与字幕文件 {timeline_file} 对应的视频文件 {video_file}")
            continue
        
        # 根据文件名创建文件夹
        folder_name = os.path.splitext(os.path.basename(timeline_file))[0]
        output_dir_qian = os.path.join(os.path.dirname(timeline_file), folder_name + "_Aqian")
        output_dir_hou = os.path.join(os.path.dirname(timeline_file), folder_name + "_Bhou")
        
        os.makedirs(output_dir_qian, exist_ok=True)
        os.makedirs(output_dir_hou, exist_ok=True)
        
        timestamps = extract_timestamps(timeline_file)
        qian_count, hou_count = capture_frames(video_file, timestamps, output_dir_qian, output_dir_hou, video_index, total_videos)
        
        total_qian_count += qian_count
        total_hou_count += hou_count
        
        video_results.append(f"第{video_index}视频处理完成!\n总共提取了 {qian_count + hou_count} 张截图\nAqian 文件夹: {qian_count} 张\nBhou 文件夹: {hou_count} 张")
    
    print("=" * 50)
    print("所有视频处理结果:")
    for result in video_results:
        print(result)
        print("-" * 30)
    
    print("=" * 50)
    print(f"所有视频处理完成!")
    print(f"总共处理了 {total_videos} 个视频")
    print(f"总共提取了 {total_qian_count + total_hou_count} 张截图")
    print(f"Aqian 文件夹总共: {total_qian_count} 张")
    print(f"Bhou 文件夹总共: {total_hou_count} 张")
    
    input("按回车键退出...")
