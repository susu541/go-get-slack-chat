package main

import (
	"bufio"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"os"
	"strconv"
	"strings"
	"time"
	_ "time/tzdata"

	"github.com/wesovilabs/koazee"
	"github.com/xuri/excelize/v2"
)

const (
	slackApiBaseUrl = "https://slack.com/api/"
)

type SlackMessagebk struct {
	User string `json:"user"`
	Text string `json:"text"`
}

type SlackMessage struct {
	Type       string `json:"type"`
	User       string `json:"user"`
	Text       string `json:"text"`
	Ts         string `json:"ts"`
	ThreadTs   string `json:"thread_ts"`
	ReplyCount int    `json:"reply_count"`
}

type ResponseMetadata struct {
	Cursor string `json:"next_cursor"`
}

type ReqParam struct {
	Key   string `json:"id"`
	Value string `json:"name"`
}

type SlackConversationsHistoryResponse struct {
	Ok       bool             `json:"ok"`
	Messages []SlackMessage   `json:"messages"`
	Error    string           `json:"error"`
	Needed   string           `json:"needed"`
	Metadata ResponseMetadata `json:"response_metadata"`
}

type Members struct {
	Id       string `json:"id"`
	Name     string `json:"name"`
	RealName string `json:"real_name"`
}

type SlackConversationsUsersResponse struct {
	Ok      bool      `json:"ok"`
	Members []Members `json:"members"`
	Error   string    `json:"error"`
	Needed  string    `json:"needed"`
}

type UserName struct {
	Name string `json:"name"`
}

type UserNamesResponse struct {
	UserName []UserName `json:"data"`
}

type config struct {
	SlackBotToken  string `json:'slackBotToken'`
	SlackUserToken string `json:'slackUserToken'`
	Channeld       string `json:'channeld'`
	GsUrl          string `json:'gsUrl'`
	GsAccessToken  string `json:'gsAccessToken'`
}

var slackBotToken string
var slackUserToken string
var gsAccessToken string
var gsUrl string

func main() {
	loc := time.FixedZone("Asia/Tokyo", 9*60*60)
	time.Local = loc
	cfg, _ := loadConfig()
	slackBotToken = cfg.SlackBotToken
	slackUserToken = cfg.SlackUserToken
	gsAccessToken = cfg.GsAccessToken
	gsUrl = cfg.GsUrl
	target := readString("月初から指定された日までの履歴を取得します。yyyy-MM-dd形式で日付を入力してください。\n")

	const layout = "2006-01-02"
	t, err := time.Parse(layout, target)
	if err != nil {
		panic(err)
	}

	userNameRes := getUserName(t)
	if len(userNameRes) < 1 {
		fmt.Println("Error get UserName:", err)
		return
	}

	members, err := getMembers()
	if err != nil {
		fmt.Println("Error get member:", err)
		return
	}

	start := time.Date(t.Year(), t.Month(), 1, 0, 0, 0, 0, loc).AddDate(0, 0, -1)
	end := time.Date(t.Year(), t.Month(), t.Day(), 23, 59, 59, 59, loc)
	messages, err := fetchSlackMessages(cfg.Channeld, t, start, end)
	if err != nil {
		fmt.Println("Error fetching slack message:", err)
		return
	}

	// 取得したメッセージをユーザごとにグループ化
	stream := koazee.StreamOf(messages)
	out, _ := stream.Sort(func(person, otherPerson SlackMessage) int {
		return strings.Compare(person.Ts, otherPerson.Ts)
	}).GroupBy(func(val SlackMessage) string { return val.User })
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	var sheetName string
	iter := out.MapRange() // return *reflect.MapIter
	for iter.Next() {
		sheetName = ""
		for _, v := range members {
			if v.Id == iter.Key().String() {
				for i := 0; i < len(userNameRes); i++ {
					if strings.HasPrefix(v.RealName, userNameRes[i].Name) {
						sheetName = v.RealName
					}
				}
			}
		}
		if len(sheetName) < 1 {
			continue
		}
		// ワークシートを作成する
		_, err := f.NewSheet(sheetName)
		if err != nil {
			fmt.Println(err)
			return
		}

		var slice []string
		var yyyymmdd []string
		var hhmmss []string
		for i := 0; i < iter.Value().Len(); i++ {
			// メッセージを取得
			msg := iter.Value().Index(i).FieldByName("Ts").String()

			// Unix 時刻を日時に変換
			ts := msg[:len(msg)-7]
			c, _ := strconv.ParseInt(ts, 10, 64)
			dtFromUnix := time.Unix(c, 0)
			if dtFromUnix.Before(start) {
				continue
			}
			yyyymmdd = append(yyyymmdd, dtFromUnix.Format("2006-01-02"))
			hhmmss = append(hhmmss, dtFromUnix.Format("15:04:05"))
			slice = append(slice, iter.Value().Index(i).FieldByName("Text").String())
		}
		// セルの値を設定
		f.SetSheetCol(sheetName, "A1", &yyyymmdd)
		f.SetSheetCol(sheetName, "B1", &hhmmss)
		f.SetSheetCol(sheetName, "C1", &slice)
		f.SetColWidth(sheetName, "A", "H", 11)
	}

	// 指定されたパスに従ってファイルを保存します
	if err := f.SaveAs(t.Format("200601") + ".xlsx"); err != nil {
		fmt.Println(err)
	}
}

// 設定ファイルの読み込み
func loadConfig() (*config, error) {
	f, err := os.Open("config.json")
	if err != nil {
		fmt.Println("loadConfig os.Open err:", err)
		return nil, err
	}
	defer f.Close()

	var cfg config
	err = json.NewDecoder(f).Decode(&cfg)
	return &cfg, err
}

// 入力文字を取得
func readString(msg string) string {
	fmt.Print(msg)
	reader := bufio.NewReader(os.Stdin)
	OutputBucket, _ := reader.ReadString('\n')

	return strings.TrimRight(OutputBucket, "\r\n")
}

// 取得対象のユーザ名一覧をスプレッドシートから取得
func getUserName(t time.Time) []UserName {
	yyyyMM := t.Format("200601")

	url := fmt.Sprintf(gsUrl, yyyyMM, gsAccessToken)
	req, err := http.NewRequest("GET", url, nil)
	if err != nil {
		return nil
	}

	client := &http.Client{Timeout: time.Second * 30}
	resp, err := client.Do(req)
	if err != nil {
		return nil
	}

	defer resp.Body.Close()

	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return nil
	}

	var apiResponse UserNamesResponse
	err = json.Unmarshal(body, &apiResponse)
	if err != nil {
		fmt.Println(err)
		return nil
	}
	return apiResponse.UserName
}

// Slackの対象ワークスペース内のユーザをすべて取得
func getMembers() ([]Members, error) {
	url := fmt.Sprintf("%susers.list", slackApiBaseUrl)
	body, err := exec("GET", "application/x-www-form-urlencoded", url, slackBotToken)
	if err != nil {
		return nil, err
	}

	var apiResponse SlackConversationsUsersResponse
	err = json.Unmarshal(body, &apiResponse)
	if err != nil {
		return nil, err
	}
	return apiResponse.Members, nil
}

// 指定したチャンネル・期間のチャットをすべて取得する
func fetchSlackMessages(channelId string, target time.Time, start time.Time, end time.Time) ([]SlackMessage, error) {
	var cursor string
	var allMsg []SlackMessage

	u, err := url.Parse(slackApiBaseUrl + "conversations.history")
	if err != nil {
		// handle error
	}
	params := []*ReqParam{
		{"channel", channelId},
		{"inclusive", "true"},
		{"include_all_metadata", "true"},
		{"limit", "100"},
		{"start", strconv.FormatInt(start.Unix(), 10)},
		{"end", strconv.FormatInt(end.Unix(), 10)},
	}

	q := u.Query()
	for _, p := range params {
		q.Set(p.Key, p.Value)
	}

	u.RawQuery = q.Encode()

	for {
		var c string
		var msg []SlackMessage
		if len(cursor) > 0 {
			q.Set("cursor", cursor)
			u.RawQuery = q.Encode()
			msg, c, _ = exeSlackApi(channelId, start, end, target, u.String())
		} else {
			msg, c, _ = exeSlackApi(channelId, start, end, target, u.String())
		}
		cursor = c

		if len(cursor) < 1 {
			break
		}

		allMsg = append(allMsg, msg...)
	}

	return allMsg, nil
}

// SlackAPIの取得
func exeSlackApi(channelId string, start time.Time, end time.Time, target time.Time, url string) ([]SlackMessage, string, error) {
	// slack api呼び出し
	body, err := exec("GET", "application/json", url, slackUserToken)
	if err != nil {
		return nil, "", err
	}

	var apiResponse SlackConversationsHistoryResponse
	err = json.Unmarshal(body, &apiResponse)
	if err != nil {
		return nil, "", err
	}
	if !apiResponse.Ok {
		return nil, "", fmt.Errorf("slack API error: %s, needed: %s", apiResponse.Error, apiResponse.Needed)
	}

	return apiResponse.Messages, apiResponse.Metadata.Cursor, nil
}

// APIの実行
func exec(method string, contentType string, url string, token string) ([]byte, error) {
	req, err := http.NewRequest(method, url, nil)
	if err != nil {
		return nil, err
	}

	req.Header.Set("Content-Type", contentType)
	req.Header.Set("Authorization", fmt.Sprintf("Bearer %s", token))

	client := &http.Client{Timeout: time.Second * 10}
	resp, err := client.Do(req)
	if err != nil {
		return nil, err
	}

	defer resp.Body.Close()

	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return nil, err
	}

	return body, nil
}
