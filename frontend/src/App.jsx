const API = import.meta.env.VITE_API_URL || "http://127.0.0.1:8000";

import { useState } from "react";
import axios from "axios";
import {
  Activity, Search, BarChart3, MessageSquare,
  TrendingUp, Users, LoaderCircle, AlertCircle,
  Heart, Menu, Sparkles
} from "lucide-react";
import {
  PieChart, Pie, Cell, Tooltip, ResponsiveContainer,
  BarChart, Bar, XAxis, YAxis, CartesianGrid
} from "recharts";
import "./App.css";

const COLORS = ["#16a34a", "#94a3b8", "#dc2626"];
const CATEGORY_COLORS = [
  "#6366f1", "#06b6d4", "#f59e0b", "#10b981",
  "#ec4899", "#8b5cf6", "#64748b", "#f97316"
];

function getUsername(value) {
  return value.trim().replace(/\/+$/, "").split("/").pop().replace(/^@/, "");
}

function StatCard({ title, value, icon: Icon, color }) {
  return (
    <div className="stat-card">
      <div>
        <p>{title}</p>
        <h2>{value}</h2>
      </div>
      <div className={`stat-icon ${color}`}>
        <Icon size={22} />
      </div>
    </div>
  );
}

export default function App() {
  const [activeTab, setActiveTab] = useState("profile");
  const [username, setUsername] = useState("");
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [topic, setTopic] = useState("");
  const [topicData, setTopicData] = useState(null);
  const [topicLoading, setTopicLoading] = useState(false);
  const [topicError, setTopicError] = useState("");

  async function analyzeProfile(e) {
    e.preventDefault();

    const user = getUsername(username);

    if (!user) {
      setError("Please enter a valid X username or profile URL.");
      return;
    }

    setLoading(true);
    setError("");
    setData(null);

    try {
      const response = await axios.get(
        `${API}/analyze/${encodeURIComponent(user)}`,
        { params: { limit: 100 }, timeout: 120000 }
      );

      setData(response.data);
    } catch (err) {
      setError(
        err.response?.data?.detail ||
        "Unable to analyze this profile. Please try again."
      );
    } finally {
      setLoading(false);
    }
  }
  async function searchTopic(e) {
  e.preventDefault();

  if (!topic.trim()) {
    setTopicError("Please enter a topic.");
    return;
  }

  setTopicLoading(true);
  setTopicError("");
  setTopicData(null);

  try {
    const response = await axios.get(
      `${API}/topic-search/${encodeURIComponent(topic.trim())}`,
      {
        params: { pages: 3 },
        timeout: 60000
      }
    );

    setTopicData(response.data);
  } catch (err) {
    setTopicError(
      err.response?.data?.detail ||
      "Unable to search this topic. Please try again."
    );
  } finally {
    setTopicLoading(false);
  }
}

  const sentiment = data?.sentiment_distribution || {};
  const categories = data?.category_distribution || {};

  const sentimentData = [
    { name: "Positive", value: sentiment.Positive || 0 },
    { name: "Neutral", value: sentiment.Neutral || 0 },
    { name: "Negative", value: sentiment.Negative || 0 }
  ];

  const categoryData = Object.entries(categories)
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => b.value - a.value);

  const total = data?.analyzed_count || 0;

  return (
    <div className="app">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-icon"><Activity /></div>
          <div>
            <h2>TwitterSense</h2>
            <span>AI INTELLIGENCE</span>
          </div>
        </div>

        <p className="nav-heading">WORKSPACE</p>

        <button
          className={`nav-item ${activeTab === "profile" ? "active" : ""}`}
          onClick={() => setActiveTab("profile")}
        >
          <BarChart3 size={19} /> Profile Analysis
        </button>

        <button
          className={`nav-item ${activeTab === "topic" ? "active" : ""}`}
          onClick={() => setActiveTab("topic")}
        >
          <Search size={19} /> Topic Intelligence
        </button>

        <div className="sidebar-bottom">
          <div className="status-dot"></div>
          Backend connection
          <span>API :8000</span>
        </div>
      </aside>

      <main className="main">
        <header className="topbar">
          <div>
            <p className="eyebrow">SOCIAL MEDIA ANALYTICS</p>
            <h1>
              {activeTab === "profile"
                ? "Profile Analysis"
                : "Topic Intelligence"}
            </h1>
          </div>
          <div className="online-badge">
            <span></span> System Online
          </div>
        </header>

        {activeTab === "profile" ? (
          <>
            <section className="welcome">
              <div>
                <h2>Understand the conversation.</h2>
                <p>
                  Analyze public posts, discover sentiment patterns
                  and identify the topics that matter.
                </p>
              </div>
              <Sparkles size={40} />
            </section>

            <section className="search-panel">
              <form onSubmit={analyzeProfile}>
                <label>X PROFILE URL OR USERNAME</label>
                <div className="search-row">
                  <div className="search-input">
                    <Search size={19} />
                    <input
                      value={username}
                      onChange={(e) => setUsername(e.target.value)}
                      placeholder="Enter username or profile URL"
                    />
                  </div>
                  <button className="primary-btn" disabled={loading}>
                    {loading ? (
                      <><LoaderCircle className="spin" size={18} /> Analyzing</>
                    ) : (
                      <><Activity size={18} /> Analyze Profile</>
                    )}
                  </button>
                </div>
              </form>
              <p className="hint">
                Up to 100 recent public posts will be analyzed.
              </p>
            </section>

            {error && (
              <div className="error-box">
                <AlertCircle size={19} /> {String(error)}
              </div>
            )}

            {loading && (
              <div className="loading-panel">
                <LoaderCircle className="spin" size={32} />
                <h3>Analyzing profile...</h3>
                <p>Fetching posts and calculating insights.</p>
              </div>
            )}

            {!data && !loading && !error && (
              <div className="empty-state">
                <div className="empty-icon"><BarChart3 size={32} /></div>
                <h3>Your analytics will appear here</h3>
                <p>Enter an X username above and click Analyze Profile.</p>
              </div>
            )}

            {data && (
              <>
                <section className="profile-heading">
                  <div className="avatar">
                    {(data.name || data.username || "X")[0].toUpperCase()}
                  </div>
                  <div>
                    <h2>{data.name || data.username}</h2>
                    <p>@{data.username}</p>
                  </div>
                  <div className="analyzed-tag">
                    <Activity size={15} /> Analysis complete
                  </div>
                </section>

                <section className="stats-grid">
                  <StatCard title="Posts Fetched" value={data.fetched_count} icon={MessageSquare} color="blue" />
                  <StatCard title="Posts Analyzed" value={total} icon={BarChart3} color="purple" />
                  <StatCard title="Positive Posts" value={sentiment.Positive || 0} icon={Heart} color="green" />
                  <StatCard title="Categories Found" value={categoryData.length} icon={TrendingUp} color="orange" />
                </section>

                
<section className="charts-grid">

  {/* Sentiment Analysis - Horizontal Bar Chart */}
  <div className="panel">
    <div className="panel-heading">
      <div>
        <h3>Sentiment Analysis</h3>
        <p>Emotional distribution of posts</p>
      </div>
    </div>

    <div className="category-chart">
      <ResponsiveContainer width="100%" height="100%">
        <BarChart
          data={sentimentData}
          layout="vertical"
          margin={{ left: 15, right: 25 }}
        >
          <CartesianGrid strokeDasharray="3 3" horizontal={false} />
          <XAxis type="number" />
          <YAxis
            dataKey="name"
            type="category"
            width={85}
            tick={{ fontSize: 12 }}
          />
          <Tooltip />
          <Bar dataKey="value" radius={[0, 7, 7, 0]}>
            {sentimentData.map((item, index) => (
              <Cell key={item.name} fill={COLORS[index]} />
            ))}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  </div>

  {/* Category Distribution - Circular Doughnut Chart */}
  <div className="panel">
    <div className="panel-heading">
      <div>
        <h3>Category Distribution</h3>
        <p>Most discussed topics</p>
      </div>
    </div>

    <div className="chart-area">
      <ResponsiveContainer width="100%" height="100%">
        <PieChart>
          <Pie
            data={categoryData}
            dataKey="value"
            nameKey="name"
            innerRadius={65}
            outerRadius={100}
            paddingAngle={4}
          >
            {categoryData.map((item, index) => (
              <Cell
                key={item.name}
                fill={CATEGORY_COLORS[index % CATEGORY_COLORS.length]}
              />
            ))}
          </Pie>
          <Tooltip />
        </PieChart>
      </ResponsiveContainer>

      <div className="chart-center">
        <strong>
          {categoryData.reduce((sum, item) => sum + item.value, 0)}
        </strong>
        <span>Posts</span>
      </div>
    </div>

    <div className="legend category-legend">
      {categoryData.map((item, index) => (
        <div key={item.name}>
          <span
            style={{
              background: CATEGORY_COLORS[index % CATEGORY_COLORS.length]
            }}
          />
          {item.name}
          <strong>{item.value}</strong>
        </div>
      ))}
    </div>
  </div>

</section>


                <section className="panel posts-panel">
                  <div className="panel-heading">
                    <div>
                      <h3>Analyzed Posts</h3>
                      <p>{total} posts classified by sentiment and category</p>
                    </div>
                    <span className="count-tag">{total} POSTS</span>
                  </div>

                  <div className="posts-list">
                    {(data.analyzed_posts || []).slice(0, 20).map((post, i) => (
                      <article className="post" key={i}>
                        <p>{post.text}</p>
                        <div className="post-tags">
                          <span className="category-tag">{post.category}</span>
                          <span className={`sentiment-tag ${post.sentiment.toLowerCase()}`}>
                            {post.sentiment}
                          </span>
                          <span className="score">
                            Score: {post.sentiment_score}
                          </span>
                        </div>
                      </article>
                    ))}
                  </div>
                </section>
              </>
            )}
          </>
        ) : (
  <>
    <section className="welcome">
      <div>
        <h2>Discover the voices behind a topic.</h2>
        <p>
          Find accounts discussing a subject and understand
          their activity and sentiment.
        </p>
      </div>
      <Users size={40} />
    </section>

    <section className="search-panel">
      <form onSubmit={searchTopic}>
        <label>ENTER A TOPIC OR KEYWORD</label>
        <div className="search-row">
          <div className="search-input">
            <Search size={19} />
            <input
              value={topic}
              onChange={(e) => setTopic(e.target.value)}
              placeholder="e.g. Artificial Intelligence"
            />
          </div>
          <button className="primary-btn" disabled={topicLoading}>
            {topicLoading ? (
              <><LoaderCircle className="spin" size={18} /> Searching</>
            ) : (
              <><Search size={18} /> Discover Accounts</>
            )}
          </button>
        </div>
        <p className="hint">
          Searches recent public posts and ranks up to 10 accounts.
        </p>
      </form>
    </section>

    {topicError && (
      <div className="error-box">
        <AlertCircle size={19} /> {String(topicError)}
      </div>
    )}

    {topicLoading && (
      <div className="loading-panel">
        <LoaderCircle className="spin" size={32} />
        <h3>Searching public conversations...</h3>
        <p>Identifying accounts discussing your topic.</p>
      </div>
    )}

    {!topicData && !topicLoading && !topicError && (
      <div className="empty-state">
        <div className="empty-icon"><Users size={32} /></div>
        <h3>Discover relevant accounts</h3>
        <p>Enter a topic and click Discover Accounts.</p>
      </div>
    )}

    {topicData && (
      <>
        <section className="stats-grid">
          <StatCard
            title="Posts Retrieved"
            value={topicData.total_posts_retrieved}
            icon={MessageSquare}
            color="blue"
          />
          <StatCard
            title="Unique Accounts"
            value={topicData.unique_accounts}
            icon={Users}
            color="purple"
          />
          <StatCard
            title="Pages Searched"
            value={topicData.pages_fetched}
            icon={Search}
            color="orange"
          />
          <StatCard
            title="Top Accounts"
            value={topicData.top_accounts?.length || 0}
            icon={TrendingUp}
            color="green"
          />
        </section>

        
<section className="panel accounts-panel">
  <div className="panel-heading">
    <div>
      <h3>Top 10 Accounts</h3>
      <p>Accounts discussing "{topicData.topic}"</p>
    </div>
    <span className="count-tag">RANKED RESULTS</span>
  </div>

  <p className="hint">
    Ranking: matching post count, followed by follower count.
  </p>

  <div className="account-table-wrapper">
    <table className="account-table">
      <thead>
        <tr>
          <th>RANK</th>
          <th>ACCOUNT</th>
          <th>FOLLOWERS</th>
          <th>POSTS</th>
          <th>SENTIMENT DISTRIBUTION</th>
        </tr>
      </thead>

      <tbody>
        {(topicData.top_accounts || []).map((account, index) => (
          <tr key={account.username}>
            <td>
              <span className="rank-badge">
                #{index + 1}
              </span>
            </td>

            <td>
              <div className="account-identity">
                <strong>{account.name}</strong>
                <small>@{account.username}</small>
              </div>
            </td>

            <td className="followers-cell">
              {Number(account.followers || 0).toLocaleString()}
            </td>

            <td>
              <span className="posts-count">
                {account.post_count}
              </span>
            </td>

            <td>
              <div className="sentiment-distribution">
                <span className="sentiment-pill positive">
                  <span className="sentiment-dot"></span>
                  {account.positive}
                </span>

                <span className="sentiment-pill neutral">
                  <span className="sentiment-dot"></span>
                  {account.neutral}
                </span>

                <span className="sentiment-pill negative">
                  <span className="sentiment-dot"></span>
                  {account.negative}
                </span>
              </div>
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  </div>
</section>

<section className="panel posts-panel">
  <div className="panel-heading">
    <div>
      <h3>Recent Tweets</h3>
      <p>Posts discussing "{topicData.topic}"</p>
    </div>
    <span className="count-tag">
      {topicData.matching_posts?.length || 0} TWEETS
    </span>
  </div>

  {(topicData.matching_posts || []).map((post, index) => (
    <article className="post" key={index}>
      <div className="tweet-author">
        <div className="tweet-avatar">
          {(post.name || post.username || "X")[0].toUpperCase()}
        </div>
        <div>
          <strong>{post.name || post.username}</strong>
          <small>@{post.username}</small>
        </div>
      </div>

      <p>{post.text}</p>

      <div className="post-tags">
        <span className="category-tag">{post.category}</span>
        <span className={`sentiment-tag ${post.sentiment.toLowerCase()}`}>
          {post.sentiment}
        </span>
        {post.url && (
          <a href={post.url} target="_blank" rel="noreferrer">
            View original ↗
          </a>
        )}
      </div>
    </article>
  ))}
</section>

      </>
    )}
  </>
)}

        <footer>TwitterSense AI · Social Media Intelligence Platform</footer>
      </main>
    </div>
  );
}
